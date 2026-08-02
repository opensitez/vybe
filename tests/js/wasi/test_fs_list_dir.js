// vybe-test: js/wasi/test_fs_list_dir
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

fs.mkdir("/tmp/vybe_test_dir");
        fs.writeFile("/tmp/vybe_test_dir/a.txt", "a");
        fs.writeFile("/tmp/vybe_test_dir/b.txt", "b");
        let files = fs.listDir("/tmp/vybe_test_dir");
        __check(__line(files.length), "2");
        fs.remove("/tmp/vybe_test_dir");
