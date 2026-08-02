// vybe-test: js/private_brand_errors/private_in_check_false_on_different_class_instance
// origin: languages/js/tests/js/test_private_brand_errors.rs

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

class A{#a=1; static isA(v){return #a in v;}} class B{#b=1;} __check(__line(A.isA(new B())), "false");
