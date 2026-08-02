// vybe-test: js/new_target_private_brand/new_target_flows_through_super_constructor
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

class Base {
  constructor() {
    __check(__line(new.target === Derived), "true");
  }
}
class Derived extends Base {}
new Derived();
