// vybe-test: js/async_generators/for_await_of_with_break
// origin: languages/js/tests/js/test_async_generators.rs

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

async function* naturals() {
    let n = 1;
    while (true) yield n++;
}
async function main() {
    const results = [];
    for await (const v of naturals()) {
        if (v > 5) break;
        results.push(v);
    }
    console.log(results.join(","));
}
main();
