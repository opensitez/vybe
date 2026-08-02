// vybe-test: js/reflect_apply_construct_get_set/test_js_reflect_apply_basic_invocation
// origin: languages/js/tests/js/test_js_reflect_apply_construct_get_set.rs

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

function add(a, b) { return a + b + this.bonus; }
const ctx = { bonus: 10 };
__check(__line(Reflect.apply(add, ctx, [5, 15])), "30");
