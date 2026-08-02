// vybe-test: js/map_set_deep/map_foreach_this_arg_binding_context
// origin: languages/js/tests/js/test_map_set_deep.rs

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

const ctx = { factor: 10 };
const m = new Map([["a", 2]]);
m.forEach(function(v, k) {
    console.log(k + ":" + (v * this.factor));
}, ctx);
