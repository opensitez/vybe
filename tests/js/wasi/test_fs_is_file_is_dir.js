// vybe-test: js/wasi/test_fs_is_file_is_dir
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

fs.mkdir("/tmp/vybe_test_isdir");
        fs.writeFile("/tmp/vybe_test_isdir/f.txt", "x");
        __check(__line(fs.isDir("/tmp/vybe_test_isdir"), fs.isFile("/tmp/vybe_test_isdir/f.txt")), "true true");
        fs.remove("/tmp/vybe_test_isdir");
