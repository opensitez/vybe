// vybe-test: js/async_patterns/async_class_method
// origin: languages/js/tests/js/test_async_patterns.rs

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

class Api {
    async fetch(id) {
        let result = await Promise.resolve("item_" + id);
        return result;
    }
}
async function main() {
    let api = new Api();
    let r = await api.fetch(42);
    console.log(r);
}
main();
