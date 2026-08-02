// vybe-test: js/throw_in_switch/switch_throw_typeerror_subclass
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

let o=[];try{switch(0){case 0:throw new TypeError("bad op");default:o.push("d");}}catch(e){o.push(e.name+":"+e.message);}__check(__line(o.join(",")), "TypeError:bad op");
