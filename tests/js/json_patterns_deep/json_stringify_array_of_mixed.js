// vybe-test: js/json_patterns_deep/json_stringify_array_of_mixed
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

const arr = [1, "two", null, true, { x: 3 }];
const json = JSON.stringify(arr);
__check(__line(json), "[1,\"two\",null,true,{\"x\":3}]");
