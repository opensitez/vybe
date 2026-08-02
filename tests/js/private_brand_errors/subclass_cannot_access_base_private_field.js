// vybe-test: js/private_brand_errors/subclass_cannot_access_base_private_field
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

class B{#secret=1;} class D extends B{read(){try{return this.#secret;}catch(e){return "err";}}} __check(__line(new D().read()), "err");
