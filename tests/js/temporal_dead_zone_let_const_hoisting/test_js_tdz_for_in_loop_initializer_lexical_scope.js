// vybe-test: js/temporal_dead_zone_let_const_hoisting/test_js_tdz_for_in_loop_initializer_lexical_scope
// origin: languages/js/tests/js/test_js_temporal_dead_zone_let_const_hoisting.rs

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

const obj = { a: 1, b: 2 };
const keys = [];
for (const k in obj) {
    keys.push(k);
}
console.log(keys.join(","));
