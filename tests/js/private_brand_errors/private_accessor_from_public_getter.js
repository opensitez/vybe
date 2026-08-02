// vybe-test: js/private_brand_errors/private_accessor_from_public_getter
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

class C{#v=2; get value(){return this.#v;} } __check(__line(new C().value), "2");
