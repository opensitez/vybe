// vybe-test: js/throw_in_switch/switch_finally_runs_after_case_throw
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

let o=[];try{try{switch(1){case 1:throw new Error("t");}}finally{o.push("f");}}catch(e){o.push(e.message);}__check(__line(o.join(",")), "f,t");
