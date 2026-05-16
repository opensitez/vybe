use super::helpers::run_js;

fn run_js_one(code: &str) -> String {
    run_js(code).into_iter().next().unwrap_or_default()
}

// ============================================================
// vybe:fs — filesystem
// ============================================================

#[test]
fn test_fs_write_read() {
    let code = r#"
        fs.writeFile("/tmp/vybe_test_fs.txt", "hello vybe");
        let content = fs.readFile("/tmp/vybe_test_fs.txt");
        console.log(content);
        fs.remove("/tmp/vybe_test_fs.txt");
    "#;
    assert_eq!(run_js_one(code), "hello vybe");
}

#[test]
fn test_fs_exists() {
    let code = r#"
        fs.writeFile("/tmp/vybe_test_exists.txt", "test");
        console.log(fs.exists("/tmp/vybe_test_exists.txt"));
        fs.remove("/tmp/vybe_test_exists.txt");
        console.log(fs.exists("/tmp/vybe_test_exists.txt"));
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "true");
    assert_eq!(lines[1], "false");
}

#[test]
fn test_fs_append() {
    let code = r#"
        fs.writeFile("/tmp/vybe_test_append.txt", "hello");
        fs.appendFile("/tmp/vybe_test_append.txt", " world");
        console.log(fs.readFile("/tmp/vybe_test_append.txt"));
        fs.remove("/tmp/vybe_test_append.txt");
    "#;
    assert_eq!(run_js_one(code), "hello world");
}

#[test]
fn test_fs_list_dir() {
    let code = r#"
        fs.mkdir("/tmp/vybe_test_dir");
        fs.writeFile("/tmp/vybe_test_dir/a.txt", "a");
        fs.writeFile("/tmp/vybe_test_dir/b.txt", "b");
        let files = fs.listDir("/tmp/vybe_test_dir");
        console.log(files.length);
        fs.remove("/tmp/vybe_test_dir");
    "#;
    assert_eq!(run_js_one(code), "2");
}

#[test]
fn test_fs_is_file_is_dir() {
    let code = r#"
        fs.mkdir("/tmp/vybe_test_isdir");
        fs.writeFile("/tmp/vybe_test_isdir/f.txt", "x");
        console.log(fs.isDir("/tmp/vybe_test_isdir"), fs.isFile("/tmp/vybe_test_isdir/f.txt"));
        fs.remove("/tmp/vybe_test_isdir");
    "#;
    assert_eq!(run_js_one(code), "true true");
}

// ============================================================
// vybe:clock — time
// ============================================================

#[test]
fn test_clock_now() {
    let code = r#"
        let t = clock.now();
        console.log(t > 0);
    "#;
    assert_eq!(run_js_one(code), "true");
}

#[test]
fn test_clock_date_now() {
    // JS alias: Date.now() → vybe:clock/now
    let code = r#"
        let t = Date.now();
        console.log(t > 1000000000000);
    "#;
    assert_eq!(run_js_one(code), "true");
}

#[test]
fn test_clock_to_iso_string() {
    let code = r#"
        let iso = clock.toISOString(0);
        console.log(iso);
    "#;
    assert_eq!(run_js_one(code), "1970-01-01T00:00:00.000Z");
}

// ============================================================
// vybe:env — environment
// ============================================================

#[test]
fn test_env_platform() {
    let code = "console.log(env.platform())";
    let result = run_js_one(code);
    assert!(result == "macos" || result == "linux" || result == "windows");
}

#[test]
fn test_env_cwd() {
    let code = "console.log(env.cwd() !== null)";
    assert_eq!(run_js_one(code), "true");
}

#[test]
fn test_env_args() {
    let code = "console.log(env.args().length > 0)";
    assert_eq!(run_js_one(code), "true");
}

// ============================================================
// vybe:random
// ============================================================

#[test]
fn test_random_random() {
    let code = r#"
        let r = random.random();
        console.log(r >= 0, r < 1);
    "#;
    assert_eq!(run_js_one(code), "true true");
}

#[test]
fn test_random_int() {
    let code = r#"
        let r = random.randomInt(1, 10);
        console.log(r >= 1, r <= 10);
    "#;
    assert_eq!(run_js_one(code), "true true");
}

#[test]
fn test_random_uuid() {
    let code = r#"
        let id = random.uuid();
        console.log(id.length);
    "#;
    assert_eq!(run_js_one(code), "36");
}

#[test]
fn test_math_random_uses_vybe_random() {
    // Math.random() should now use the proper PRNG
    let code = r#"
        let r = Math.random();
        console.log(r >= 0, r < 1);
    "#;
    assert_eq!(run_js_one(code), "true true");
}

// ============================================================
// vybe:http (basic test — uses localhost to avoid network deps)
// ============================================================

#[test]
fn test_http_fetch_error() {
    // Connecting to a closed port should return an error
    let code = r#"
        let resp = http.fetch("http://127.0.0.1:1");
        console.log(resp.ok);
    "#;
    assert_eq!(run_js_one(code), "false");
}

// ============================================================
// Combined: real-world patterns
// ============================================================

#[test]
fn test_write_json_read_back() {
    let code = r#"
        let data = { name: "Alice", age: 30 };
        fs.writeFile("/tmp/vybe_test_json.json", JSON.stringify(data));
        let raw = fs.readFile("/tmp/vybe_test_json.json");
        console.log(raw);
        fs.remove("/tmp/vybe_test_json.json");
    "#;
    let result = run_js_one(code);
    // JSON key order isn't guaranteed, so check both fields are present
    assert!(result.contains("\"name\":\"Alice\""));
    assert!(result.contains("\"age\":30"));
}

#[test]
fn test_timed_operation() {
    let code = r#"
        let start = clock.now();
        let sum = 0;
        for (let i = 0; i < 10000; i++) { sum = sum + i; }
        let elapsed = clock.now() - start;
        console.log(sum, elapsed >= 0);
    "#;
    assert_eq!(run_js_one(code), "49995000 true");
}

// ============================================================
// JSON.parse
// ============================================================

#[test]
fn test_json_parse_object() {
    let code = r#"
        let obj = JSON.parse('{"name":"Alice","age":30}');
        console.log(obj.name, obj.age);
    "#;
    assert_eq!(run_js_one(code), "Alice 30");
}

#[test]
fn test_json_parse_array() {
    let code = r#"
        let arr = JSON.parse("[1, 2, 3]");
        console.log(arr.length, arr[0]);
    "#;
    assert_eq!(run_js_one(code), "3 1");
}

#[test]
fn test_json_roundtrip() {
    let code = r#"
        let obj = { x: 10, y: "hello" };
        let json = JSON.stringify(obj);
        let parsed = JSON.parse(json);
        console.log(parsed.x, parsed.y);
    "#;
    assert_eq!(run_js_one(code), "10 hello");
}

// ============================================================
// Base64
// ============================================================

#[test]
fn test_btoa_atob() {
    let code = r#"
        let encoded = btoa("Hello World");
        console.log(encoded);
        let decoded = atob(encoded);
        console.log(decoded);
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "SGVsbG8gV29ybGQ=");
    assert_eq!(lines[1], "Hello World");
}

// ============================================================
// URI encoding
// ============================================================

#[test]
fn test_encode_uri_component() {
    assert_eq!(run_js_one(r#"console.log(encodeURIComponent("hello world"))"#), "hello%20world");
}

#[test]
fn test_decode_uri_component() {
    assert_eq!(run_js_one(r#"console.log(decodeURIComponent("hello%20world"))"#), "hello world");
}

// ============================================================
// Array.isArray / Array.from
// ============================================================

#[test]
fn test_array_is_array() {
    let code = r#"
        console.log(Array.isArray([1, 2, 3]));
        console.log(Array.isArray("hello"));
        console.log(Array.isArray(42));
    "#;
    let lines = run_js(code);
    assert_eq!(lines[0], "true");
    assert_eq!(lines[1], "false");
    assert_eq!(lines[2], "false");
}

// ============================================================
// wasi:filesystem stat/rename/copy
// ============================================================

#[test]
fn test_fs_stat() {
    let code = r#"
        fs.writeFile("/tmp/vybe_stat_test.txt", "hello");
        let info = fs.stat("/tmp/vybe_stat_test.txt");
        console.log(info.isFile, info.size);
        fs.remove("/tmp/vybe_stat_test.txt");
    "#;
    assert_eq!(run_js_one(code), "true 5");
}

#[test]
fn test_fs_rename() {
    let code = r#"
        fs.writeFile("/tmp/vybe_rename_old.txt", "data");
        fs.rename("/tmp/vybe_rename_old.txt", "/tmp/vybe_rename_new.txt");
        console.log(fs.exists("/tmp/vybe_rename_old.txt"), fs.exists("/tmp/vybe_rename_new.txt"));
        fs.remove("/tmp/vybe_rename_new.txt");
    "#;
    assert_eq!(run_js_one(code), "false true");
}

#[test]
fn test_fs_copy() {
    let code = r#"
        fs.writeFile("/tmp/vybe_copy_src.txt", "original");
        fs.copy("/tmp/vybe_copy_src.txt", "/tmp/vybe_copy_dst.txt");
        console.log(fs.readFile("/tmp/vybe_copy_dst.txt"));
        fs.remove("/tmp/vybe_copy_src.txt");
        fs.remove("/tmp/vybe_copy_dst.txt");
    "#;
    assert_eq!(run_js_one(code), "original");
}
