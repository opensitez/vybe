// vybe-test: js/json_serialization/json_nested_array_object
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

const data = JSON.parse('[{"id":1,"tags":["a","b"]},{"id":2,"tags":["c"]}]');
__check(__line(data.length), "2");
__check(__line(data[0].tags.join(",")), "a,b");
__check(__line(data[1].id), "2");
