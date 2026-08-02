// vybe-test: js/class_private_methods_and_getters/test_js_class_private_generator_method
// origin: languages/js/tests/js/test_js_class_private_methods_and_getters.rs

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

class Sequence {
    async *#generate() {
        yield 1;
        yield 2;
    }
    async getItems() {
        const res = [];
        for await (const x of this.#generate()) res.push(x);
        return res.join(",");
    }
}
new Sequence().getItems().then(res => console.log(res));
