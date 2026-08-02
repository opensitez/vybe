// vybe-test: js/private_brand_errors/private_method_called_with_super_in_subclass_public
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

class B{#base(){return "b";} pub(){return this.#base();}} class D extends B{wrap(){return super.pub();}} __check(__line(new D().wrap()), "b");
