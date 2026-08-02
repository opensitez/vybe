// vybe-test: js/explicit_resource_management/using_in_for_of_disposes_each_iteration
// origin: languages/js/tests/js/test_explicit_resource_management.rs

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

const log = [];
function makeRes(n) {
    return { n, [Symbol.dispose]() { log.push("d" + this.n); } };
}
for (using r of [makeRes(1), makeRes(2), makeRes(3)]) {
    log.push("u" + r.n);
}
console.log(log.join(","));
