// vybe-test: js/throw_in_switch/switch_multiple_labels_share_throw_body
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

let o=[];try{switch(2){case 1:case 2:throw new Error("shared");default:o.push("d");}}catch(e){o.push(e.message);}__check(__line(o.join(",")), "shared");
