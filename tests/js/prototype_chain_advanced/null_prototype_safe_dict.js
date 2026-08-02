// vybe-test: js/prototype_chain_advanced/null_prototype_safe_dict
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

// Use null-prototype objects as safe dicts (no prototype pollution)
const dict = Object.create(null);
dict.constructor = "overridden"; // doesn't affect anything
dict.hasOwnProperty = "also overridden";
console.log(dict.constructor);
console.log(Object.hasOwn(dict, "constructor"));
