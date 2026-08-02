// vybe-test: js/json_patterns_deep/json_custom_tojson
// origin: languages/js/tests/js/test_json_patterns_deep.rs

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

class Temperature {
    constructor(celsius) { this.celsius = celsius; }
    toJSON() {
        return { celsius: this.celsius, fahrenheit: this.celsius * 9/5 + 32 };
    }
}
const t = new Temperature(100);
const json = JSON.stringify(t);
const parsed = JSON.parse(json);
__check(__line(parsed.celsius), "100");
__check(__line(parsed.fahrenheit), "212");
