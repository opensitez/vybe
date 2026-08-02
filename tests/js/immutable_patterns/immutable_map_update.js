// vybe-test: js/immutable_patterns/immutable_map_update
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

const config = new Map([["theme", "light"], ["lang", "en"]]);
// Create new Map with update
const updated = new Map([...config, ["theme", "dark"]]);
__check(__line(config.get("theme")), "light");
__check(__line(updated.get("theme")), "dark");
__check(__line(updated.get("lang")), "en");
