// vybe-test: js/object_literal_advanced/async_method_shorthand
// origin: languages/js/tests/js/test_object_literal_advanced.rs

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
    async fetchData() {
        const v = await Promise.resolve(42);
        return v;
    }
};
async function main() {
    console.log(await obj.fetchData());
}
main();
