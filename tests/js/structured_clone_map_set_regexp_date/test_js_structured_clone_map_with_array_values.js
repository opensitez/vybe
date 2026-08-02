// vybe-test: js/structured_clone_map_set_regexp_date/test_js_structured_clone_map_with_array_values
// origin: languages/js/tests/js/test_js_structured_clone_map_set_regexp_date.rs

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

const map = new Map([["nums", [1, 2, 3]]]);
const clone = structuredClone(map);
clone.get("nums").push(4);
__check(__line(map.get("nums").length + "|" + clone.get("nums").length), "3|4");
