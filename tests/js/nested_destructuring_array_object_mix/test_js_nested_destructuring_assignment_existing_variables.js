// vybe-test: js/nested_destructuring_array_object_mix/test_js_nested_destructuring_assignment_existing_variables
// origin: languages/js/tests/js/test_js_nested_destructuring_array_object_mix.rs

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

let a, b, c;
[{ a, inner: [b, c] }] = [{ a: 1, inner: [2, 3] }];
__check(__line(`${a},${b},${c}`), "1,2,3");
