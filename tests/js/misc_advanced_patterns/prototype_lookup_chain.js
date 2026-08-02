// vybe-test: js/misc_advanced_patterns/prototype_lookup_chain
// origin: languages/js/tests/js/test_misc_advanced_patterns.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

const base = { type: "base", describe() { return this.type + ":" + this.name; } };
const mid = Object.create(base);
mid.type = "mid";
const leaf = Object.create(mid);
leaf.name = "leaf";
__check(__line(leaf.describe()), "mid:leaf");
__check(__line(leaf.type), "mid");
__check(__line(leaf.hasOwnProperty("name")), "true");
__check(__line(leaf.hasOwnProperty("type")), "false");
