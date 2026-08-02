// vybe-test: js/new_target_private_brand/new_target_is_undefined_in_plain_call
// origin: languages/js/tests/js/test_new_target_private_brand.rs

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

function kind() {
  return new.target === undefined ? "plain" : "construct";
}
__check(__line(kind()), "plain");
