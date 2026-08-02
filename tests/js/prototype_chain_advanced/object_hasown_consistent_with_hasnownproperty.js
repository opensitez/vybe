// vybe-test: js/prototype_chain_advanced/object_hasown_consistent_with_hasnownproperty
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

const obj = Object.create({ inherited: true });
obj.own = true;
__check(__line(obj.hasOwnProperty("own")), "true");
__check(__line(obj.hasOwnProperty("inherited")), "false");
__check(__line(Object.hasOwn(obj, "own")), "true");
__check(__line(Object.hasOwn(obj, "inherited")), "false");
