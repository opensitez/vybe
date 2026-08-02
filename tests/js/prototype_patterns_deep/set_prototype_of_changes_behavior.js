// vybe-test: js/prototype_patterns_deep/set_prototype_of_changes_behavior
// origin: languages/js/tests/js/test_prototype_patterns_deep.rs

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

const a = { whoAmI() { return "A"; } };
const b = { whoAmI() { return "B"; } };
const obj = Object.create(a);
__check(__line(obj.whoAmI()), "A");
Object.setPrototypeOf(obj, b);
__check(__line(obj.whoAmI()), "B");
