// vybe-test: js/new_target_private_brand/new_target_is_captured_by_arrow_in_constructor
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

function Box() {
  const read = () => new.target;
  __check(__line(read() === Box), "true");
}
new Box();
