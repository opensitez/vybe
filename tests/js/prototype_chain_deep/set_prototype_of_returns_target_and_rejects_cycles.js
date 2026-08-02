// vybe-test: js/prototype_chain_deep/set_prototype_of_returns_target_and_rejects_cycles
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

const root = {};
const mid = Object.create(root);
const unchanged = Object.setPrototypeOf(root, Object.getPrototypeOf(root)) === root;
let cycleError = false;
try {
    Object.setPrototypeOf(root, mid);
} catch (e) {
    cycleError = e instanceof TypeError;
}
__check(__line(unchanged), "true");
__check(__line(cycleError), "true");
