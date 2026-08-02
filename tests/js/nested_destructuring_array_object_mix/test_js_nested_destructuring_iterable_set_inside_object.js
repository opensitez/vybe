// vybe-test: js/nested_destructuring_array_object_mix/test_js_nested_destructuring_iterable_set_inside_object
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

const obj = { numbers: new Set([100, 200]) };
const { numbers: [n1, n2] } = obj;
__check(__line(`${n1}:${n2}`), "100:200");
