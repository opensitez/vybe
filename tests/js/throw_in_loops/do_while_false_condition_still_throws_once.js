// vybe-test: js/throw_in_loops/do_while_false_condition_still_throws_once
// origin: languages/js/tests/js/test_throw_in_loops.rs

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

let o=[];let run=true;do{try{if(run)throw new Error("once");}catch(e){o.push(e.message);run=false;}o.push("body");}while(false);console.log(o.join(","));
