// vybe-test: js/throw_in_loops/nested_loop_inner_throw_caught_by_outer_try
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

let o=[];try{for(let i=0;i<2;i++){for(let j=0;j<2;j++){if(i===1&&j===1)throw new Error("deep");o.push(i+""+j);}}}catch(e){o.push(e.message);}console.log(o.join(","));
