// vybe-test: js/intl_ecma402_surface/test_js_intl_collator_surface_options

function assert(cond, msg) {
    if (!cond) {
        throw new Error(msg);
    }
}

const collator = new Intl.Collator("en-US", {
    numeric: true,
    sensitivity: "base",
    ignorePunctuation: true,
    caseFirst: "upper",
    collation: "emoji",
});
const opts = collator.resolvedOptions();
assert(opts.numeric === true, "numeric resolved option");
assert(opts.sensitivity === "base", "sensitivity resolved option");
assert(opts.ignorePunctuation === true, "ignorePunctuation resolved option");
assert(opts.caseFirst === "upper", "caseFirst resolved option");
assert(opts.collation === "emoji", "collation resolved option");
assert(collator.compare("red, envelope", "red envelope") === 0, "ignore punctuation compare");
assert(["a", "A", "b", "B"].sort(collator.compare).join("") === "AaBb", "caseFirst upper sort");

let threw = false;
try {
    collator.compare(Symbol("a"), "a");
} catch (e) {
    threw = true;
}
assert(threw, "Symbol compare throws");
console.log("ok");
