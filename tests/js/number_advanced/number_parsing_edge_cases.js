// vybe-test: js/number_advanced/number_parsing_edge_cases
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

__check(__line(parseInt("0x1F")), "31");
__check(__line(parseInt("077")), "77");
__check(__line(parseInt("3.99")), "3");
__check(__line(parseFloat("3.14abc")), "3.14");
__check(__line(Number("  42  ")), "42");
__check(__line(Number("")), "0");
