// vybe-test: js/throw_in_switch/switch_labeled_break_exits_before_later_throw
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

let o=[];outer:for(let i=0;i<2;i++){try{switch(i){case 0:o.push("ok");break outer;case 1:throw new Error("late");}}catch(e){o.push("c");}}console.log(o.join(","));
