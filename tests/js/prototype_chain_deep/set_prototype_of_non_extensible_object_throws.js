// vybe-test: js/prototype_chain_deep/set_prototype_of_non_extensible_object_throws
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

const fixed = Object.preventExtensions({});
let threw = false;
try {
    Object.setPrototypeOf(fixed, {});
} catch (e) {
    threw = e instanceof TypeError;
}
__check(__line(threw), "true");
