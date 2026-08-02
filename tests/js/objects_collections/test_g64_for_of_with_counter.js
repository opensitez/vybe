// vybe-test: js/objects_collections/test_g64_for_of_with_counter
// origin: languages/js/tests/js/js_objects_collections_test.rs

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

let arr = ["x", "y", "z"];
        let result = [];
        let idx = 0;
        for (let item of arr) {
            result.push(idx + ":" + item);
            idx = idx + 1;
        }
        console.log(result.join(","));
