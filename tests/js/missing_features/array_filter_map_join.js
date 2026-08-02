// vybe-test: js/missing_features/array_filter_map_join
// origin: languages/js/tests/js/js_missing_features_test.rs

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

let result = [1,2,3,4,5]
            .filter(x => x % 2 === 1)
            .map(x => x * 10)
            .join(",");
        __check(__line(result), "10,30,50");
