// vybe-test: js/proxy_reflect/reflect_apply_calling_a_function
// origin: languages/js/tests/js/test_proxy_reflect.rs

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

function add(a, b) { return a + b; }
// Reflect.apply with null this
const result = Reflect.apply(add, null, [3, 4]);
__check(__line(result), "7");
