// vybe-test: js/misc_es_features/function_name_property
// origin: languages/js/tests/js/test_misc_es_features.rs

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

function myFunc() {}
const arrow = () => {};
const obj = { method() {} };
__check(__line(myFunc.name), "myFunc");
__check(__line(arrow.name), "arrow");
__check(__line(obj.method.name), "method");
