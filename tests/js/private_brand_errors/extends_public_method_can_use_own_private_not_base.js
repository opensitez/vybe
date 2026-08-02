// vybe-test: js/private_brand_errors/extends_public_method_can_use_own_private_not_base
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

class B{#b=1; getB(){return this.#b;}} class D extends B{getD(){return this.#d;} #d=2;} const d=new D();__check(__line(d.getD()), "2");__check(__line(d.getB()), "1");
