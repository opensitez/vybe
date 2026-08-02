// vybe-test: js/async_await_deep/async_class_method
// origin: languages/js/tests/js/test_async_await_deep.rs

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

class DataService {
    async fetch(id) {
        return await Promise.resolve({ id, name: "Item " + id });
    }
}
async function main() {
    const svc = new DataService();
    const item = await svc.fetch(5);
    console.log(item.id);
    console.log(item.name);
}
main();
