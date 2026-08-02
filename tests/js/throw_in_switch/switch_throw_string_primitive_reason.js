// vybe-test: js/throw_in_switch/switch_throw_string_primitive_reason
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

let o=[];try{switch("x"){case "x":throw "str";default:o.push("d");}}catch(e){o.push(typeof e+":"+e);}__check(__line(o.join(",")), "string:str");
