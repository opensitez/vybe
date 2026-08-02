// vybe-test: js/throw_in_switch/switch_inside_for_loop_throw_stops_iteration_in_catch
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

let o=[];for(let i=0;i<3;i++){try{switch(i){case 1:throw new Error("stop");default:o.push(i);}}catch(e){o.push("c");break;}}console.log(o.join(","));
