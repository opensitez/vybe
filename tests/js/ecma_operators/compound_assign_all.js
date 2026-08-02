// vybe-test: js/ecma_operators/compound_assign_all
// origin: languages/js/tests/js/test_ecma_operators.rs

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

let x = 10;
x += 5; __check(__line(x), "15");
x -= 3; __check(__line(x), "12");
x *= 2; __check(__line(x), "24");
x /= 4; __check(__line(x), "6");
x %= 5; __check(__line(x), "1");
