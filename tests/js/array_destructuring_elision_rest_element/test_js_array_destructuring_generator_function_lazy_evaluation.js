// vybe-test: js/array_destructuring_elision_rest_element/test_js_array_destructuring_generator_function_lazy_evaluation
// origin: languages/js/tests/js/test_js_array_destructuring_elision_rest_element.rs

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

let evaluated = 0;
function* gen() {
    evaluated++; yield 1;
    evaluated++; yield 2;
    evaluated++; yield 3;
}
const [a, b] = gen();
__check(__line(`${a},${b}|evaluated=${evaluated}`), "1,2|evaluated=2");
