// vybe-test: js/prototype_chain_advanced/set_prototype_of_changes_chain
// origin: languages/js/tests/js/test_prototype_chain_advanced.rs

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

const a = { hello() { return "from a"; } };
const b = { hello() { return "from b"; } };
const obj = Object.create(a);
__check(__line(obj.hello()), "from a");
Object.setPrototypeOf(obj, b);
__check(__line(obj.hello()), "from b");
