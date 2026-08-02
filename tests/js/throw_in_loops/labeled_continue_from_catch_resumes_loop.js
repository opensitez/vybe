// vybe-test: js/throw_in_loops/labeled_continue_from_catch_resumes_loop
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

let o=[];outer:for(let i=0;i<3;i++){try{if(i===1)throw "x";o.push(i);}catch{continue outer;}o.push("a");}console.log(o.join(","));
