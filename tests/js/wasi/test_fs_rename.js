// vybe-test: js/wasi/test_fs_rename
// origin: languages/js/tests/js/js_wasi_test.rs

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

fs.writeFile("/tmp/vybe_rename_old.txt", "data");
        fs.rename("/tmp/vybe_rename_old.txt", "/tmp/vybe_rename_new.txt");
        __check(__line(fs.exists("/tmp/vybe_rename_old.txt"), fs.exists("/tmp/vybe_rename_new.txt")), "false true");
        fs.remove("/tmp/vybe_rename_new.txt");
