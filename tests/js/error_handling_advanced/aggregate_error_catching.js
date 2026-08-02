// vybe-test: js/error_handling_advanced/aggregate_error_catching
// origin: languages/js/tests/js/test_error_handling_advanced.rs

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
    try {
        await Promise.any([
            Promise.reject(new Error("e1")),
            Promise.reject(new Error("e2")),
        ]);
    } catch(e) {
        console.log(e instanceof AggregateError);
        console.log(e.errors.length);
        console.log(e.errors[0].message);
    }
}
main();
