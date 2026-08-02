// vybe-test: js/async_patterns/async_static_method
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

class Factory {
    static async create(name) {
        let data = await Promise.resolve({ name });
        return data;
    }
}
async function main() {
    let obj = await Factory.create("test");
    console.log(obj.name);
}
main();
