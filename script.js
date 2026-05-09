
let excelData = [];

// ========================================
// LOGIN
// ========================================

function checkLogin() {

    const login = document.getElementById("login").value;
    const password = document.getElementById("password").value;

    if (login === "admin" && password === "admin") {

        document.getElementById("loginOverlay").style.display = "none";

    } else {

        document.getElementById("loginError").innerText = "Неверный логин или пароль";

    }
}


// ========================================
// START LOGIN SCREEN
// ========================================

window.addEventListener("load", () => {
    document.getElementById("loginOverlay").style.display = "flex";
});


// ========================================
// LOAD EXCEL
// ========================================

async function loadExcel(fileName) {

    try {

        const response = await fetch(fileName);

        const arrayBuffer = await response.arrayBuffer();

        const workbook = XLSX.read(arrayBuffer, {
            type: "array"
        });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        excelData = XLSX.utils.sheet_to_json(sheet);

        console.log("Excel loaded:", excelData);

        calculatePrice();

    } catch (error) {

        console.error("Excel load error:", error);

    }
}

// старт
loadExcel("prices.xlsx");


// ========================================
// ЮР / ФИЗ (ТОЛЬКО ОДИН)
// ========================================

document.querySelectorAll(".client-only").forEach((checkbox) => {

    checkbox.addEventListener("change", function () {

        if (this.checked) {

            document.querySelectorAll(".client-only").forEach((item) => {
                if (item !== this) item.checked = false;
            });

        }

        calculatePrice();

    });

});


// ========================================
// СКИДКИ (ТОЛЬКО ОДНА)
// ========================================

document.querySelectorAll(".exclusive-option").forEach((checkbox) => {

    checkbox.addEventListener("change", function () {

        if (this.checked) {

            document.querySelectorAll(".exclusive-option").forEach((item) => {
                if (item !== this) item.checked = false;
            });

        }

        calculatePrice();

    });

});


// ========================================
// ОБЩИЕ СОБЫТИЯ
// ========================================

document.querySelectorAll("input, select").forEach((el) => {
    el.addEventListener("change", calculatePrice);
});


// ========================================
// РАСЧЁТ ЦЕНЫ
// ========================================

function calculatePrice() {

    if (!excelData.length) return;

    const zone = document.getElementById("zone").value;

    const vehicleType = document.getElementById("vehicleType").value;

    const taxi = document.getElementById("taxi").checked ? "Так" : "Ні";


    // CLIENT TYPE (строго один)
    let clientType = "";

    if (document.getElementById("jur").checked) {
        clientType = "Юридична особа";
    }

    if (document.getElementById("fiz").checked) {
        clientType = "Фізична особа";
    }


    const discount =
        document.querySelector(".exclusive-option:checked")?.value || "";


    const row = excelData.find((item) => {

        return (
            String(item.zone).trim() === String(zone).trim() &&
            String(item.vehicleType).trim() === String(vehicleType).trim() &&
            String(item.taxi).trim() === String(taxi).trim() &&
            String(item.clientType).trim() === String(clientType).trim()
        );

    });


    const priceElement = document.getElementById("price");

    if (!row) {
        priceElement.innerHTML = "Немає ціни";
        return;
    }

    let price = Number(row.price);

    if (discount === "09") price *= 0.9;
    if (discount === "095") price *= 0.95;

    priceElement.innerHTML = price.toFixed(2) + " грн";
}