// vybe-test: js/generators_advanced/async_generator_method_in_object_literal
// origin: languages/js/tests/js/test_generators_advanced.rs

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

const obj = {
    async *gen() {
        yield await Promise.resolve("async_val");
    }
};
(async () => {
    const a = [];
    for await (const v of obj.gen()) a.push(v);
    console.log(a.join(","));
})();
