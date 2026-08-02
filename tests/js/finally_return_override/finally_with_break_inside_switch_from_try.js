// vybe-test: js/finally_return_override/finally_with_break_inside_switch_from_try
// origin: languages/js/tests/js/test_finally_return_override.rs

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

let o=[];try{switch(1){case 1:try{throw "x";}finally{o.push("f");break;}default:o.push("d");}}catch(e){o.push(String(e));}__check(__line(o.join(",")), "f");
