// vybe-test: js/map_set_iterator_more_matrix/map_for_each_third_argument_is_map
// origin: languages/js/tests/js/test_map_set_iterator_more_matrix.rs

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

const m = new Map([["a", 1]]);
let ok = false;
m.forEach((_, __, self) => { ok = self === m; });
console.log(ok);
