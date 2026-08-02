// vybe-test: js/destructuring_comprehensive/default_value_is_expression
// origin: languages/js/tests/js/test_destructuring_comprehensive.rs

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

let counter = 0;
function inc() { return ++counter; }
const { a = inc(), b = inc(), c = 5 } = { b: 99 };
__check(__line(a), "1");
__check(__line(b), "99");
__check(__line(c), "5");
__check(__line(counter), "1");
