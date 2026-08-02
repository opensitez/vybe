// vybe-test: js/prototype_chain_deep/object_create_chain_of_three_levels
// origin: languages/js/tests/js/test_prototype_chain_deep.rs

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

const a = { level: "a" };
const b = Object.create(a);
b.bProp = "b";
const c = Object.create(b);
__check(__line(c.level), "a");
__check(__line(c.bProp), "b");
__check(__line(Object.getPrototypeOf(Object.getPrototypeOf(c)) === a), "true");
