// vybe-test: js/json_serialization/json_reviver_transform
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

const json = '{"createdAt":"2024-01-15T10:00:00.000Z","amount":"42.5"}';
const parsed = JSON.parse(json, (key, val) => {
    if (key === "createdAt") return new Date(val).getFullYear();
    if (key === "amount") return parseFloat(val);
    return val;
});
__check(__line(parsed.createdAt), "2024");
__check(__line(parsed.amount), "42.5");
__check(__line(typeof parsed.amount), "number");
