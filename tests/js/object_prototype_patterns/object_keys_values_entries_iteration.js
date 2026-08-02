// vybe-test: js/object_prototype_patterns/object_keys_values_entries_iteration
// origin: languages/js/tests/js/test_object_prototype_patterns.rs

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

const config = { host: "localhost", port: 8080, debug: true };
const pairs = Object.entries(config)
    .map(([k, v]) => k + "=" + v)
    .join(",");
__check(__line(pairs), "host=localhost,port=8080,debug=true");
