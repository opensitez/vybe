// vybe-test: js/control_flow_advanced/switch_falls_through_cases_without_break
// origin: languages/js/tests/js/test_control_flow_advanced.rs

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

let result = [];
switch (1) {
    case 1: result.push("one");
    case 2: result.push("two");
    case 3: result.push("three"); break;
    case 4: result.push("four");
}
__check(__line(result.join(",")), "one,two,three");
