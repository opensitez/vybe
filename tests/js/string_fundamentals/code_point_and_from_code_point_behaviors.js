// vybe-test: js/string_fundamentals/code_point_and_from_code_point_behaviors
// origin: languages/js/tests/js/test_string_fundamentals.rs

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

__check(__line(String.fromCodePoint(0x1F600).length), "2");
__check(__line(String.fromCodePoint(0x1F600).charCodeAt(0)), "55357");
__check(__line(String.fromCodePoint(0x1F600).charCodeAt(1)), "56832");
__check(__line(String.fromCharCode(0x41).length), "1");
