// vybe-test: js/prototype_chain_deep/is_prototype_of_follows_chain
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

const grandparent = { level: "grandparent" };
const parent = Object.create(grandparent);
const child = Object.create(parent);

__check(__line(grandparent.isPrototypeOf(child)), "true");
__check(__line(parent.isPrototypeOf(child)), "true");
__check(__line(Object.prototype.isPrototypeOf(child)), "true");
__check(__line(grandparent.isPrototypeOf(parent)), "true");
