// vybe-test: js/json_semantics_matrix/json_stringify_preserves_insertion_order_for_strings
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
obj.z = 1;
obj.a = 2;
obj.m = 3;
__check(__line(JSON.stringify(obj)), "{\"z\":1,\"a\":2,\"m\":3}");
