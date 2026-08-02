// vybe-test: js/immutable_patterns/immutable_update_nested
// origin: languages/js/tests/js/test_immutable_patterns.rs

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

const state = Object.freeze({
    user: { name: "Bob", settings: { theme: "light" } },
    count: 0
});
const newState = {
    ...state,
    user: { ...state.user, settings: { ...state.user.settings, theme: "dark" } },
    count: state.count + 1
};
__check(__line(state.user.settings.theme), "light");
__check(__line(newState.user.settings.theme), "dark");
__check(__line(newState.count), "1");
