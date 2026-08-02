// vybe-test: js/getter_setter_deep/read_only_getter
// origin: languages/js/tests/js/test_getter_setter_deep.rs

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

const obj = {
    get pi() { return 3.14159; }
};
obj.pi = 99; // silently ignored in non-strict
__check(__line(obj.pi), "3.14159");
