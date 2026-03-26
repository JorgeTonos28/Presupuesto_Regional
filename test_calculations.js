const assert = require('assert');

function round2_(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

const baseHoras = 10;
const baseAcciones = 5;
const basePresupuesto = 1000;
const multiplier = 0.17;

const rowAccionesScaled = baseAcciones * multiplier;
const rowPresupuestoScaled = basePresupuesto * multiplier;
const rowTotalHoras = baseHoras * rowAccionesScaled;
const rowCostoAF = rowTotalHoras > 0 ? (rowPresupuestoScaled / rowTotalHoras) : 0;

console.log("Acciones Escalonadas:", rowAccionesScaled); // 0.85
console.log("Presupuesto Escalonado:", rowPresupuestoScaled); // 170
console.log("Total Horas:", rowTotalHoras); // 8.5
console.log("Costo AF:", rowCostoAF); // 20

assert.strictEqual(rowTotalHoras, 8.5);
assert.strictEqual(rowCostoAF, 20);

const testBaseRegional = 746407015.00;
const testPct17 = 17;
const testPct15 = 15;
const testPct20 = 20;
const testPct48 = 48;

const testAmount17 = round2_(testBaseRegional * (testPct17 / 100));
const testAmount15 = round2_(testBaseRegional * (testPct15 / 100));
const testAmount20 = round2_(testBaseRegional * (testPct20 / 100));
const testAmount48 = round2_(testBaseRegional * (testPct48 / 100));

console.log("17% segment:", testAmount17);
console.log("15% segment:", testAmount15);
console.log("20% segment:", testAmount20);
console.log("48% segment:", testAmount48);

assert.strictEqual(testAmount17, 126889192.55);
assert.strictEqual(testAmount15, 111961052.25);
assert.strictEqual(testAmount20, 149281403.00);
assert.strictEqual(testAmount48, 358275367.20);

console.log("Tests OK");