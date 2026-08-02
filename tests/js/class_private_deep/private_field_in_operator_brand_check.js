// vybe-test: js/class_private_deep/private_field_in_operator_brand_check
// origin: languages/js/tests/js/test_class_private_deep.rs

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

class Foo {
    #x;
    static isFoo(obj) { return #x in obj; }
}
const f = new Foo();
__check(__line(Foo.isFoo(f)), "true");
__check(__line(Foo.isFoo({})), "false");
