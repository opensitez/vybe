// vybe-test: js/array_higher_order/array_methods_skip_missing_holes
// origin: languages/js/tests/js/test_array_higher_order.rs

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

const sparse = [, 2, , 4];
const doubled = sparse.map(x => x * 2);
__check(__line(doubled.length), "4");
__check(__line(0 in doubled, 1 in doubled, 2 in doubled), "false true false");
__check(__line(doubled.join("|")), "|4||8");
