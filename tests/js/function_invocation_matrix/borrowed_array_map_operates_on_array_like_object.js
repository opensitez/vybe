// vybe-test: js/function_invocation_matrix/borrowed_array_map_operates_on_array_like_object
// origin: languages/js/tests/js/test_function_invocation_matrix.rs

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

const arrLike = { 0: 1, 1: 2, length: 2 };
const out = Array.prototype.map.call(arrLike, x => x * 3);
__check(__line(out.join(",")), "3,6");
