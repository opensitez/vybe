// vybe-test: js/promise_all_allsettled_race_any_combinators/test_js_promise_all_sparse_array_preserves_holes
// origin: languages/js/tests/js/test_js_promise_all_allsettled_race_any_combinators.rs

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

const arr = [Promise.resolve(1), , Promise.resolve(3)];
Promise.all(arr).then(res => console.log(res[0] + "|" + res[1] + "|" + res[2]));
