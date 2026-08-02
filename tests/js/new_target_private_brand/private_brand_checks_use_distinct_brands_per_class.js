// vybe-test: js/new_target_private_brand/private_brand_checks_use_distinct_brands_per_class
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

class Left {
  #value = 1;
  hasBrand(value) {
    return #value in value;
  }
}
class Right {
  #value = 1;
}
__check(__line(new Left().hasBrand(new Right())), "false");
