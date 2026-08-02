// vybe-test: js/template_literal_advanced/template_substitutions_evaluate_left_to_right
// origin: languages/js/tests/js/test_template_literal_advanced.rs

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

const events = [];
function mark(label) {
    events.push(label);
    return label;
}
__check(__line(`${mark("first")}-${mark("second")}-${mark("third")}`), "first-second-third");
__check(__line(events.join(",")), "first,second,third");
