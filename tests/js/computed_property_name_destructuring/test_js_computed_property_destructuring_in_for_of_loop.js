// vybe-test: js/computed_property_name_destructuring/test_js_computed_property_destructuring_in_for_of_loop
// origin: languages/js/tests/js/test_js_computed_property_name_destructuring.rs

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

const records = [{ k: "a", a: 1 }, { k: "b", b: 2 }];
const results = [];
for (const { k, [k]: val } of records) {
    results.push(val);
}
console.log(results.join(","));
