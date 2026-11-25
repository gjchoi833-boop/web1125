"use strict";

const DAILY_AMOUNT = 25000;
const DAY_MS = 86400000;
const NIGHT_CAPS = { seoul: 100000, metro: 80000, other: 70000 };
const money = new Intl.NumberFormat("ko-KR");

function tripDays(departure, returning) {
  if (!departure || !returning) return 0;
  return (Date.parse(returning) - Date.parse(departure)) / DAY_MS + 1;
}

function localAllowance(hours, vehicle) {
  if (vehicle === "assigned") return 0;
  const amount = hours === "long" ? 20000 : 10000;
  return amount - (vehicle === "official" ? 10000 : 0);
}

function longStayReduction(days) {
  const afterArrival = days - 1;
  const days16to30 = Math.max(0, Math.min(afterArrival, 30) - 15);
  const days31to60 = Math.max(0, Math.min(afterArrival, 60) - 30);
  const daysAfter60 = Math.max(0, afterArrival - 60);
  return 2500 * days16to30 + 5000 * days31to60 + 7500 * daysAfter60;
}

function lodgingAllowance(grade, higher, nights) {
  return nights.reduce((sum, { region, count, cost }) => {
    if (grade === "1") return sum + count * cost;
    const cap = NIGHT_CAPS[region] * (higher ? 1.3 : 1);
    return sum + count * Math.min(cost, cap);
  }, 0);
}

function outsideAllowance({ days, half, mealsExcluded, longStay, grade, higher, nights, fare, mileage }) {
  const daily = DAILY_AMOUNT * (days - half / 2) - (longStay ? longStayReduction(days) : 0);
  const meals = DAILY_AMOUNT * (days - mealsExcluded);
  return [
    ["운임", fare],
    ["일비", daily],
    ...(mileage > 0 ? [["공적 마일리지 추가 일비", mileage]] : []),
    ["숙박비", lodgingAllowance(grade, higher, nights)],
    ["식비", meals]
  ];
}

if (typeof document !== "undefined") {
  const get = (id) => document.getElementById(id);
  const form = get("trip-form");
  const nightsList = get("lodging-list");
  const errorBox = get("form-error");
  const baselineRows = [["운임", null], ["일비", null], ["숙박비", null], ["식비", null]];
  const basisNote = "입력한 조건과 증빙에 따라 실제 정산액이 달라질 수 있습니다.";
  let currentResult = null;
  const selectedScope = () => form.elements.scope.value;
  const daysEntered = () => tripDays(get("departure").value, get("return").value);
  const show = (id, visible) => { get(id).hidden = !visible; };

  function showField(container, field, visible, required = false) {
    show(container, visible);
    const input = get(field);
    input.disabled = !visible;
    input.required = visible && required;
  }

  function message(error) {
    errorBox.textContent = error;
    errorBox.hidden = !error;
  }

  function renderResult(rows, context, note = basisNote) {
    const calculated = rows[0][1] !== null;
    const total = rows.reduce((sum, [, amount]) => sum + (amount ?? 0), 0);
    get("total").textContent = calculated ? money.format(total) : "—";
    get("result-context").textContent = context;
    get("result-note").textContent = note;
    currentResult = calculated
      ? { departure: get("departure").value, returning: get("return").value, rows, total, context, note,
          basis: get("basis-title").textContent }
      : null;
    get("result-dates").textContent = calculated
      ? `출장 일정 ${currentResult.departure} ~ ${currentResult.returning}` : "";
    get("result-dates").hidden = !calculated;
    get("print-result").disabled = !calculated;
    get("download-excel").disabled = !calculated;
    get("breakdown").replaceChildren(...rows.map(([label, amount]) => {
      const row = document.createElement("div");
      row.className = "result-row";
      const term = document.createElement("dt");
      const value = document.createElement("dd");
      term.textContent = label;
      value.textContent = amount === null ? "—" : `${money.format(amount)}원`;
      row.append(term, value);
      return row;
    }));
  }

  function clearResult() {
    const rows = selectedScope() === "inside" ? [["근무지 내 출장여비", null]] : baselineRows;
    renderResult(rows, "출장 정보를 입력하고 계산해 보세요.");
  }

  function renameNights() {
    Array.from(nightsList.children, (row, index) => {
      const title = `숙박 내역 ${index + 1}`;
      row.querySelector(".night-title").textContent = title;
      row.querySelector("button").setAttribute("aria-label", `${title} 삭제`);
      row.querySelector("button").hidden = nightsList.children.length === 1;
    });
  }

  function addNight(count, focus = false) {
    const row = get("night-template").content.firstElementChild.cloneNode(true);
    row.querySelector(".night-count").value = count;
    nightsList.append(row);
    renameNights();
    if (focus) row.querySelector("select").focus();
  }

  function transportDetails(grade) {
    return {
      none: ["", ""],
      rail: ["왕복 철도운임", `제${grade}호 ${grade === "1" ? "특실" : "일반실"} 기준 실비를 입력하세요.`],
      ship: ["왕복 선박운임", `제${grade}호 ${grade === "1" ? "1등급" : "2등급"} 기준 실비를 입력하세요.`],
      air: ["왕복 항공운임", "좌석 등급 기준과 공적 마일리지 사용분을 반영한 실제 지급 대상 운임을 입력하세요."],
      bus: ["왕복 버스운임", "정해진 요율 범위의 버스요금을 입력하세요."],
      car: ["왕복 기준 운임", "해당 구간의 철도운임 또는 버스운임을 입력하세요."],
      "car-exception": ["지급 대상 비용 합계", "기관 기준에 따라 인정되는 연료비·통행료 등을 입력하세요."],
      official: ["", "공용 차량·선박·항공기의 운임은 0원입니다. 공무용 차량 이용에 해당하는 일수는 일비 반액 조건에 반영하세요."],
      mixed: ["왕복 지급 대상 운임 합계", "수단별 지급 기준을 적용한 금액을 합산하고, 공용 교통수단에서 발생한 운임은 제외하세요."]
    };
  }

  function updateForm() {
    const days = daysEntered();
    const departure = get("departure").value;
    const returning = get("return").value;
    get("return").min = departure;
    get("duration").textContent = days > 0 ? `출장 ${days}일 · 날짜 간격 ${days - 1}박`
      : departure && returning ? "복귀일은 출장일과 같거나 그 이후여야 합니다."
      : "출장일과 복귀일을 선택하세요.";

    for (const scope of ["inside", "outside"]) {
      const active = selectedScope() === scope;
      show(scope, active);
      get(scope).disabled = !active;
    }

    const half = get("half-mode").value === "some";
    showField("half-field", "half-days", half, true);
    get("half-days").max = days > 0 ? Math.max(0, days - 1) : "";
    get("meal-days").max = days > 0 ? days : "";

    const transport = get("transport").value;
    const [label, note] = transportDetails(get("grade").value)[transport];
    showField("fare-field", "fare", Boolean(label), true);
    get("fare-label").textContent = label;
    get("fare-note").textContent = note;
    show("fare-note", Boolean(note));

    const lodging = get("has-lodging").checked;
    show("lodging-fields", lodging);
    get("lodging-fields").disabled = !lodging;
    if (lodging && nightsList.children.length === 0) addNight(days > 1 ? days - 1 : 1);
    const higher = lodging && get("grade").value === "2";
    showField("higher-field", "higher-lodging", higher);
    show("higher-note", higher);

    const airline = transport === "air" || transport === "mixed";
    showField("mileage-field", "mileage-amount", airline);
    showField("savings-field", "mileage-savings", airline);
    show("mileage-note", airline);
  }

  function wholeNumber(input, label) {
    const number = input.value === "" ? 0 : input.valueAsNumber;
    if (!Number.isSafeInteger(number) || number < 0) {
      throw new Error(`${label}에 0 이상의 정수를 입력하세요.`);
    }
    return number;
  }

  function readNights(days) {
    if (!get("has-lodging").checked) return [];
    const nights = Array.from(nightsList.children, (row) => ({
      region: row.querySelector(".night-region").value,
      count: wholeNumber(row.querySelector(".night-count"), "숙박 박수"),
      cost: wholeNumber(row.querySelector(".night-cost"), "1박 결제액")
    }));
    if (nights.reduce((sum, night) => sum + night.count, 0) > days - 1) {
      throw new Error("숙박 박수 합계는 출장일과 복귀일 사이의 박수를 초과할 수 없습니다.");
    }
    return nights;
  }

  function calculateOutside(days) {
    const grade = get("grade").value;
    const halfMode = get("half-mode").value;
    const half = halfMode === "all" ? days : halfMode === "some"
      ? wholeNumber(get("half-days"), "일비 반액 적용 일수") : 0;
    const mealsExcluded = wholeNumber(get("meal-days"), "식비 제외 일수");
    const fare = get("fare").disabled ? 0 : wholeNumber(get("fare"), "운임");
    const longStay = get("long-stay").checked;
    const nights = readNights(days);
    const higher = grade === "2" && nights.length > 0 && get("higher-lodging").checked;
    const mileage = get("mileage-amount").disabled ? 0 : wholeNumber(get("mileage-amount"), "추가 일비");
    const savings = get("mileage-savings").disabled ? 0 : wholeNumber(get("mileage-savings"), "절약한 항공운임");
    if (mileage > Math.min(savings, days * DAILY_AMOUNT / 2)) {
      throw new Error("추가 일비는 절약한 항공운임과 출장 기간 기본 일비의 50%를 각각 초과할 수 없습니다.");
    }
    if (days > 16 && longStay && half > 0) {
      throw new Error("장기 체재 감액과 일비 반액이 함께 적용되는 출장은 소속 기관에서 일비 적용 기준을 확인하세요.");
    }

    const rows = outsideAllowance({ days, half, mealsExcluded, longStay, grade, higher, nights, fare, mileage });
    if (!Number.isSafeInteger(rows.reduce((sum, [, amount]) => sum + amount, 0))) {
      throw new Error("입력한 금액이 너무 큽니다. 운임과 숙박비를 확인하세요.");
    }
    const paidNights = nights.reduce((sum, night) => sum + night.count, 0);
    const note = higher
      ? "국내 숙박비 초과분은 출장 종료 다음 날부터 1주 이내에 증빙과 함께 정산을 신청해야 합니다."
      : "운임·숙박비는 원칙적으로 출장 종료 다음 날부터 2주 이내에 증빙과 함께 정산을 신청합니다.";
    renderResult(rows, `근무지 외 · 제${grade}호 · ${days}일 · 숙박 청구 ${paidNights}박`, note);
  }

  form.addEventListener("input", () => {
    updateForm();
    message("");
    clearResult();
  });

  form.addEventListener("invalid", (event) => {
    const details = event.target.closest("details");
    if (details) details.open = true;
  }, true);

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    message("");
    try {
      const days = daysEntered();
      if (!Number.isInteger(days) || days < 1) throw new Error("출장일과 복귀일을 올바르게 입력하세요.");
      if (selectedScope() === "inside") {
        if (days !== 1) throw new Error("근무지 내 출장은 출장 날짜별로 시간을 입력해 각각 계산하세요.");
        const hours = get("inside-hours").value;
        const vehicle = get("inside-vehicle").value;
        renderResult([["근무지 내 출장여비", localAllowance(hours, vehicle)]],
          `근무지 내 · ${hours === "long" ? "4시간 이상" : "4시간 미만"}`);
      } else {
        calculateOutside(days);
      }
      if (window.matchMedia("(max-width: 760px)").matches) get("result-title").scrollIntoView({ block: "start" });
    } catch (error) {
      clearResult();
      message(error.message);
    }
  });

  get("print-result").addEventListener("click", () => window.print());
  for (const button of document.querySelectorAll("[data-help-dialog]")) {
    button.addEventListener("click", () => get(button.dataset.helpDialog).showModal());
  }
  get("download-excel").addEventListener("click", () => {
    if (currentResult) saveTripWorkbook(currentResult);
  });

  get("add-night").addEventListener("click", () => {
    addNight(1, true);
    message("");
    clearResult();
  });

  nightsList.addEventListener("click", (event) => {
    const remove = event.target.closest(".remove-night");
    if (!remove || nightsList.children.length <= 1) return;
    remove.closest(".night-row").remove();
    renameNights();
    get("add-night").focus();
    message("");
    clearResult();
  });

  form.addEventListener("reset", () => {
    window.requestAnimationFrame(() => {
      nightsList.replaceChildren();
      form.querySelector("details").open = false;
      message("");
      updateForm();
      clearResult();
    });
  });

  updateForm();
  clearResult();
}
