// vybe-test: js/prototype_chain_advanced/instanceof_via_symbol_hasinstance
// origin: languages/js/tests/js/test_prototype_chain_advanced.rs

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

class Range {
    constructor(min, max) { this.min = min; this.max = max; }
}
const r = new Range(0, 10);
__check(__line(r instanceof Range), "true");
__check(__line(r instanceof Object), "true");
