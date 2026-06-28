//! `System.IO.File` — read, write, append, exists, delete, lines.
use super::helpers::run_csharp;

#[test]
fn write_all_text_then_read_all_text_roundtrips() {
    assert_eq!(
        run_csharp(
            r#"
string path = System.IO.Path.GetTempFileName();
System.IO.File.WriteAllText(path, "hello");
Console.WriteLine(System.IO.File.ReadAllText(path));
System.IO.File.Delete(path);
"#
        ),
        &["hello"]
    );
}

#[test]
fn write_all_lines_then_read_all_lines_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"
string path = System.IO.Path.GetTempFileName();
System.IO.File.WriteAllLines(path, new[]{"a","b","c"});
var lines = System.IO.File.ReadAllLines(path);
Console.WriteLine(lines.Length);
Console.WriteLine(lines[1]);
System.IO.File.Delete(path);
"#
        ),
        &["3", "b"]
    );
}

#[test]
fn append_all_text_adds_to_existing_file() {
    assert_eq!(
        run_csharp(
            r#"
string path = System.IO.Path.GetTempFileName();
System.IO.File.WriteAllText(path, "hello");
System.IO.File.AppendAllText(path, " world");
Console.WriteLine(System.IO.File.ReadAllText(path));
System.IO.File.Delete(path);
"#
        ),
        &["hello world"]
    );
}

#[test]
fn file_exists_returns_false_for_nonexistent_path() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.IO.File.Exists("/no/such/path/xyz123.txt"));"#),
        &["False"]
    );
}

#[test]
fn file_exists_returns_true_after_write() {
    assert_eq!(
        run_csharp(
            r#"
string path = System.IO.Path.GetTempFileName();
Console.WriteLine(System.IO.File.Exists(path));
System.IO.File.Delete(path);
"#
        ),
        &["True"]
    );
}

#[test]
fn file_copy_produces_identical_content() {
    assert_eq!(
        run_csharp(
            r#"
string src = System.IO.Path.GetTempFileName();
string dst = src + ".copy";
System.IO.File.WriteAllText(src, "data");
System.IO.File.Copy(src, dst, true);
Console.WriteLine(System.IO.File.ReadAllText(dst));
System.IO.File.Delete(src);
System.IO.File.Delete(dst);
"#
        ),
        &["data"]
    );
}

#[test]
fn read_all_bytes_count_matches_written_byte_length() {
    assert_eq!(
        run_csharp(
            r#"
string path = System.IO.Path.GetTempFileName();
System.IO.File.WriteAllBytes(path, new byte[]{1,2,3,4,5});
Console.WriteLine(System.IO.File.ReadAllBytes(path).Length);
System.IO.File.Delete(path);
"#
        ),
        &["5"]
    );
}
