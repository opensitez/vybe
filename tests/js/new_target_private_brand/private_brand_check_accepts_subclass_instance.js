// vybe-test: js/new_target_private_brand/private_brand_check_accepts_subclass_instance
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
  #token = 1;
  hasBrand(value) {
    return #token in value;
  }
}
class Derived extends Base {}
__check(__line(new Base().hasBrand(new Derived())), "true");
