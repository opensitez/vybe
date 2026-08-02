// vybe-test: js/throw_in_switch/switch_case_block_with_local_before_throw
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

let o=[];try{switch(1){case 1:{const x=2;throw new Error("x"+x);}default:o.push("d");}}catch(e){o.push(e.message);}__check(__line(o.join(",")), "x2");
