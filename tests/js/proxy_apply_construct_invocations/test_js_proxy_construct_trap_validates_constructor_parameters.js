// vybe-test: js/proxy_apply_construct_invocations/test_js_proxy_construct_trap_validates_constructor_parameters
// origin: languages/js/tests/js/test_js_proxy_apply_construct_invocations.rs

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

class Product {
    constructor(price) { this.price = price; }
}
const ValidatedProduct = new Proxy(Product, {
    construct(target, args) {
        if (args[0] <= 0) throw new Error("Invalid Price");
        return new target(...args);
    }
});
try {
    new ValidatedProduct(-10);
} catch (e) {
    __check(__line(e.message), "Invalid Price");
}
