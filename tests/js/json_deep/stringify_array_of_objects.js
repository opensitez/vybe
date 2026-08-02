// vybe-test: js/json_deep/stringify_array_of_objects
// origin: languages/js/tests/js/test_json_deep.rs

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

const arr = [{ id: 1, name: "a" }, { id: 2, name: "b" }];
const result = JSON.parse(JSON.stringify(arr));
__check(__line(result.length), "2");
__check(__line(result[1].name), "b");
