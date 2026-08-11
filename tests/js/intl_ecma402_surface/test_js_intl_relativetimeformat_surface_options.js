// vybe-test: js/intl_ecma402_surface/test_js_intl_relativetimeformat_surface_options

function assert(cond, msg) {
    if (!cond) {
        throw new Error(msg);
    }
}

const auto = new Intl.RelativeTimeFormat("EN-us", { numeric: "auto" });
assert(auto.resolvedOptions().locale === "en-US", "canonical locale");
assert(auto.format(0, "second") === "now", "auto now");
assert(auto.format(-1, "quarter") === "last quarter", "auto last quarter");
assert(auto.format(1, "quarter") === "next quarter", "auto next quarter");

const always = new Intl.RelativeTimeFormat("en", { numeric: "always" });
const parts = always.formatToParts(-1, "day");
assert(parts.map((p) => `${p.type}:${p.value}`).join("|") === "integer:1|literal: day ago", "formatToParts structure");
assert(always.format(-0, "second") === "0 seconds ago", "negative zero direction");

let badUnit = false;
try {
    always.format(1, "fortnight");
} catch (e) {
    badUnit = true;
}
assert(badUnit, "invalid unit throws");
console.log("ok");
