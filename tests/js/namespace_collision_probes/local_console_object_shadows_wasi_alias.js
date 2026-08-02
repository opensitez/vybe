// vybe-test: js/namespace_collision_probes/local_console_object_shadows_wasi_alias
// origin: languages/js/tests/js/test_namespace_collision_probes.rs

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

{
  let console2 = { log: (x) => x };
  __check(__line(console2.log("shadow-ok")), "shadow-ok");
}
