// vybe-test: js/ecma_objects/object_spread_order_last_wins
// origin: languages/js/tests/js/test_ecma_objects.rs

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

const merged = { a: 1, ...{ a: 2, b: 3 }, a: 4 };
__check(__line(merged.a), "4");
__check(__line(merged.b), "3");
