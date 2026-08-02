// vybe-test: js/private_brand_errors/private_getter_throw_inside_accessor
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

class C{get #bad(){throw new Error("priv");} read(){try{return this.#bad;}catch(e){return e.message;}}} __check(__line(new C().read()), "priv");
