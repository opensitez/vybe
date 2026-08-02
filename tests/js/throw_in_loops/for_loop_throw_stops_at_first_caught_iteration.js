// vybe-test: js/throw_in_loops/for_loop_throw_stops_at_first_caught_iteration
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

let o=[];for(let i=0;i<4;i++){try{if(i===2)throw new Error("mid");o.push(i);}catch(e){o.push("c");break;}}console.log(o.join(","));
