// vybe-test: js/private_brand_errors/nested_class_private_not_visible_to_outer
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

class Outer{static make(){class Inner{#x=1; get(){return this.#x;}} return new Inner();}} __check(__line(Outer.make().get()), "1");
