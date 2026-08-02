// vybe-test: js/scope_prototype/let_in_for_loop_scope
// origin: languages/js/tests/js/test_scope_prototype.rs

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

let results = [];
for (let i = 0; i < 3; i++) {
    results.push(function() { return i; });
}
console.log(results[0]());
console.log(results[1]());
console.log(results[2]());
