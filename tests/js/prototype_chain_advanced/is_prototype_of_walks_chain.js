// vybe-test: js/prototype_chain_advanced/is_prototype_of_walks_chain
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

const root = { marker: "root" };
const mid = Object.create(root);
const leaf = Object.create(mid);
__check(__line(root.isPrototypeOf(leaf)), "true");
__check(__line(mid.isPrototypeOf(leaf)), "true");
__check(__line(leaf.isPrototypeOf(root)), "false");
__check(__line(Object.prototype.isPrototypeOf(leaf)), "true");
