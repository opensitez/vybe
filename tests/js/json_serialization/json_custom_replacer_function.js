// vybe-test: js/json_serialization/json_custom_replacer_function
// origin: languages/js/tests/js/test_json_serialization.rs

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

const obj = { a: 1, b: undefined, c: null, d: function(){}, e: 2 };
const json = JSON.stringify(obj, (key, val) => {
    if (val === undefined || typeof val === "function") return undefined;
    return val;
});
const parsed = JSON.parse(json);
__check(__line(parsed.a), "1");
__check(__line("b" in parsed), "false");
__check(__line("d" in parsed), "false");
__check(__line(parsed.e), "2");
