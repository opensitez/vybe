// vybe-test: js/try_catch_finally_edge/try_catch_in_loop
// origin: languages/js/tests/js/test_try_catch_finally_edge.rs

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

const results = [];
for (let i = 0; i < 3; i++) {
    try {
        if (i === 1) throw new Error("skip");
        results.push(i);
    } catch {
        results.push("err");
    }
}
console.log(results.join(","));
