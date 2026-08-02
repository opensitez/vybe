// vybe-test: js/closure_scope_deep_patterns/closure_with_generator
// origin: languages/js/tests/js/test_closure_scope_deep_patterns.rs

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

function* createIdGen(prefix) {
    let id = 0;
    while (true) yield `${prefix}-${++id}`;
}
const userIds = createIdGen("user");
const postIds = createIdGen("post");
console.log(userIds.next().value);
console.log(userIds.next().value);
console.log(postIds.next().value);
console.log(userIds.next().value);
