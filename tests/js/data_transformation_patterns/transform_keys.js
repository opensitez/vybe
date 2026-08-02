// vybe-test: js/data_transformation_patterns/transform_keys
// origin: languages/js/tests/js/test_data_transformation_patterns.rs

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

function mapKeys(obj, fn) {
    return Object.fromEntries(
        Object.entries(obj).map(([k, v]) => [fn(k), v])
    );
}
const obj = { firstName: "Alice", lastName: "Smith" };
const snaked = mapKeys(obj, k => k.replace(/([A-Z])/g, '_$1').toLowerCase());
__check(__line(snaked.first_name), "Alice");
__check(__line(snaked.last_name), "Smith");
