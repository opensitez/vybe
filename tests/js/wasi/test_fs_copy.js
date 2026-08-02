// vybe-test: js/wasi/test_fs_copy
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

fs.writeFile("/tmp/vybe_copy_src.txt", "original");
        fs.copy("/tmp/vybe_copy_src.txt", "/tmp/vybe_copy_dst.txt");
        __check(__line(fs.readFile("/tmp/vybe_copy_dst.txt")), "original");
        fs.remove("/tmp/vybe_copy_src.txt");
        fs.remove("/tmp/vybe_copy_dst.txt");
