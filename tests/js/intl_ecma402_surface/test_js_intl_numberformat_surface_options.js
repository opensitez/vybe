// vybe-test: js/intl_ecma402_surface/test_js_intl_numberformat_surface_options

function assert(cond, msg) {
    if (!cond) {
        throw new Error(msg);
    }
}

const accounting = new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    currencySign: "accounting",
});
assert(accounting.format(-100) === "($100.00)", "accounting currency sign");

const grouped = new Intl.NumberFormat("en-US", { useGrouping: false });
assert(grouped.format(1000000) === "1000000", "useGrouping false");

const signed = new Intl.NumberFormat("en-US", { signDisplay: "always" });
assert(signed.format(5) === "+5", "positive signDisplay");
assert(signed.format(0) === "+0", "zero signDisplay");

const unit = new Intl.NumberFormat("en-US", {
    style: "unit",
    unit: "meter",
    unitDisplay: "long",
});
assert(unit.format(5) === "5 meters", "long unit display");

const range = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
assert(range.formatRange(3, 5) === "$3.00 – $5.00", "currency range");
assert(range.formatRangeToParts(3, 5).some((part) => part.type === "currency"), "range parts currency");

assert(new Intl.NumberFormat("en-US").format(1000000000000000n) === "1,000,000,000,000,000", "bigint format");
console.log("ok");
