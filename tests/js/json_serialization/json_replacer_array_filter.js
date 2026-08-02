// vybe-test: js/json_serialization/json_replacer_array_filter
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

const obj = { name: "Alice", age: 30, password: "secret", email: "a@b.com" };
const json = JSON.stringify(obj, ["name", "email"]);
const parsed = JSON.parse(json);
__check(__line(parsed.name), "Alice");
__check(__line(parsed.email), "a@b.com");
__check(__line("age" in parsed), "false");
__check(__line("password" in parsed), "false");
