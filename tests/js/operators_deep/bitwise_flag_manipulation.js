// vybe-test: js/operators_deep/bitwise_flag_manipulation
// origin: languages/js/tests/js/test_operators_deep.rs

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

const READ   = 0b001;
const WRITE  = 0b010;
const EXEC   = 0b100;

let perms = READ | WRITE; // 0b011
__check(__line((perms & READ) !== 0), "true");  // has READ
__check(__line((perms & EXEC) !== 0), "false");  // has EXEC (no)
perms |= EXEC;                       // add EXEC
__check(__line((perms & EXEC) !== 0), "true");  // has EXEC (yes)
perms &= ~WRITE;                     // remove WRITE
__check(__line((perms & WRITE) !== 0), "false"); // has WRITE (no)
