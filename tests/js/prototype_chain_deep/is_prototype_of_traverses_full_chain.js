// vybe-test: js/prototype_chain_deep/is_prototype_of_traverses_full_chain
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

const a = {};
const b = Object.create(a);
const c = Object.create(b);
// Traverse chain with getPrototypeOf instead of isPrototypeOf
console.log(Object.getPrototypeOf(c) === b);
console.log(Object.getPrototypeOf(b) === a);
// a is not in c's chain going down
console.log(Object.getPrototypeOf(a) !== c);
