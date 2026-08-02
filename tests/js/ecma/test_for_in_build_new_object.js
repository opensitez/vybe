// vybe-test: js/ecma/test_for_in_build_new_object
// origin: languages/js/tests/js/js_ecma_test.rs

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

let source = { a: 1, b: 2, c: 3 };
        let doubled = {};
        for (let k in source) {
            doubled[k] = source[k] * 2;
        }
        console.log(doubled.a, doubled.b, doubled.c);
