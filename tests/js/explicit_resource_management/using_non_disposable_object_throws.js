// vybe-test: js/explicit_resource_management/using_non_disposable_object_throws
// origin: languages/js/tests/js/test_explicit_resource_management.rs

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

let threw = false;
try {
    using r = { value: 42 }; // no Symbol.dispose
} catch (e) {
    threw = e instanceof TypeError;
}
__check(__line(threw), "true");
