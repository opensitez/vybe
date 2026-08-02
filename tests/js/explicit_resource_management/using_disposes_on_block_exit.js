// vybe-test: js/explicit_resource_management/using_disposes_on_block_exit
// origin: languages/js/tests/js/test_explicit_resource_management.rs

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

const log = [];
function makeResource(name) {
    return { [Symbol.dispose]() { log.push("close:" + name); } };
}
{
    using a = makeResource("A");
    using b = makeResource("B");
    log.push("work");
}
__check(__line(log.join(",")), "work,close:B,close:A");
