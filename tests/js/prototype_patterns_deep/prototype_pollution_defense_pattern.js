// vybe-test: js/prototype_patterns_deep/prototype_pollution_defense_pattern
// origin: languages/js/tests/js/test_prototype_patterns_deep.rs

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

// Using null-prototype objects as safe maps
const safe = Object.create(null);
safe.key = "value";
__check(__line(safe.key), "value");
// No inherited methods
__check(__line(typeof safe.toString), "undefined");
__check(__line(typeof safe.hasOwnProperty), "undefined");
