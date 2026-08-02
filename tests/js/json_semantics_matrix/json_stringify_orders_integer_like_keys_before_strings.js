// vybe-test: js/json_semantics_matrix/json_stringify_orders_integer_like_keys_before_strings
// origin: languages/js/tests/js/test_json_semantics_matrix.rs

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

const obj = {};
obj.b = 1;
obj["10"] = 2;
obj.a = 3;
obj["2"] = 4;
__check(__line(JSON.stringify(obj)), "{\"2\":4,\"10\":2,\"b\":1,\"a\":3}");
