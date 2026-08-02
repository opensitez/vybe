// vybe-test: js/getter_setter_deep/lazy_initialization_getter_pattern
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

let computed = 0;
const obj = {
    get expensive() {
        // Replace with own property after first access
        const value = ++computed;
        Object.defineProperty(this, "expensive", { value, writable: true });
        return value;
    }
};
__check(__line(obj.expensive), "1"); // 1
__check(__line(obj.expensive), "1"); // 1 (cached — own property now)
__check(__line(computed), "1");       // 1
