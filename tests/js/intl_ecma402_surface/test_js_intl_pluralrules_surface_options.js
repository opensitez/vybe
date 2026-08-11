// vybe-test: js/intl_ecma402_surface/test_js_intl_pluralrules_surface_options

function assert(cond, msg) {
    if (!cond) {
        throw new Error(msg);
    }
}

const ordinal = new Intl.PluralRules("en-US", { type: "ordinal" });
assert(ordinal.resolvedOptions().pluralCategories.join(",") === "one,two,few,other", "ordinal categories");

const fractional = new Intl.PluralRules("en-US", { minimumFractionDigits: 2 });
assert(fractional.select(1) === "other", "minimumFractionDigits affects operands");

const cardinal = new Intl.PluralRules("en-US");
assert(cardinal.selectRange(0, 1) === "other", "selectRange start one");
assert(cardinal.selectRange(1, 2) === "other", "selectRange end other");

let badRange = false;
try {
    cardinal.selectRange(5, 1);
} catch (e) {
    badRange = true;
}
assert(badRange, "selectRange reversed throws");

let badSymbol = false;
try {
    cardinal.select(Symbol("1"));
} catch (e) {
    badSymbol = true;
}
assert(badSymbol, "Symbol select throws");
console.log("ok");
