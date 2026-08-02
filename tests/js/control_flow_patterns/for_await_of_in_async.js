// vybe-test: js/control_flow_patterns/for_await_of_in_async
// origin: languages/js/tests/js/test_control_flow_patterns.rs

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

async function main() {
    const promises = [1, 2, 3].map(x => Promise.resolve(x * x));
    const results = [];
    for await (const v of promises) results.push(v);
    console.log(results.join(","));
}
main();
