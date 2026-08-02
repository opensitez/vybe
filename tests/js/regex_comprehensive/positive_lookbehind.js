// vybe-test: js/regex_comprehensive/positive_lookbehind
// origin: languages/js/tests/js/test_regex_comprehensive.rs

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

const prices = "apple: $10, banana: $5, cherry: $15";
const amounts = prices.match(/(?<=\$)\d+/g);
__check(__line(amounts.join(",")), "10,5,15");
