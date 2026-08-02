// vybe-test: js/map_set_iterator_more_matrix/set_for_each_thisarg_applies_on_each_call
// origin: languages/js/tests/js/test_map_set_iterator_more_matrix.rs

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

const ctx = { base: 5 };
const out = [];
new Set([1, 2]).forEach(function(v) { out.push(v + this.base); }, ctx);
console.log(out.join(","));
