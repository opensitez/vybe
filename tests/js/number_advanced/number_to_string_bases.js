// vybe-test: js/number_advanced/number_to_string_bases
// origin: languages/js/tests/js/test_number_advanced.rs

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

const n = 255;
__check(__line(n.toString(2)), "11111111");
__check(__line(n.toString(8)), "377");
__check(__line(n.toString(16)), "ff");
__check(__line(n.toString(36)), "73");
// And parsing back
__check(__line(parseInt("ff", 16)), "255");
__check(__line(parseInt("11111111", 2)), "255");
