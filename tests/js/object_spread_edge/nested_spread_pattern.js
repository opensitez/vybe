// vybe-test: js/object_spread_edge/nested_spread_pattern
// origin: languages/js/tests/js/test_object_spread_edge.rs

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

const state = { user: { name: "Alice", age: 30 }, count: 0 };
// Update nested immutably
const next = { ...state, user: { ...state.user, age: 31 }, count: state.count + 1 };
__check(__line(next.user.name), "Alice");
__check(__line(next.user.age), "31");
__check(__line(next.count), "1");
__check(__line(state.user.age), "30"); // original unchanged
