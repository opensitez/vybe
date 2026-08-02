// vybe-test: js/throw_in_switch/switch_strict_equality_string_one_not_match_number_one
// origin: languages/js/tests/js/test_throw_in_switch.rs

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

let o=[];try{switch(1){case "1":throw new Error("str");default:throw new Error("def");}}catch(e){o.push(e.message);}__check(__line(o.join(",")), "def");
