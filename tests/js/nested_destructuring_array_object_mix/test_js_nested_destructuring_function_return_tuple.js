// vybe-test: js/nested_destructuring_array_object_mix/test_js_nested_destructuring_function_return_tuple
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

function getResponse() {
    return [200, { data: "SuccessPayload" }];
}
const [status, { data }] = getResponse();
__check(__line(`${status}:${data}`), "200:SuccessPayload");
