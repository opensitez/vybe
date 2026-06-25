//! `System.IO.Directory` — create, delete, exists, enumerate.
use super::helpers::run_csharp;

#[test]
fn directory_create_makes_new_folder() {
    assert_eq!(
        run_csharp(r#"
string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "vybe_test_"+System.Guid.NewGuid().ToString("N"));
System.IO.Directory.CreateDirectory(path);
Console.WriteLine(System.IO.Directory.Exists(path));
System.IO.Directory.Delete(path);
"#),
        &["True"]
    );
}

#[test]
fn directory_exists_returns_false_for_absent_path() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.IO.Directory.Exists("/no/such/dir/xyz999"));"#),
        &["False"]
    );
}

#[test]
fn get_files_lists_created_files_in_directory() {
    assert_eq!(
        run_csharp(r#"
string dir = System.IO.Path.Combine(System.IO.Path.GetTempPath(),"vybe_"+System.Guid.NewGuid().ToString("N"));
System.IO.Directory.CreateDirectory(dir);
System.IO.File.WriteAllText(System.IO.Path.Combine(dir,"a.txt"),"a");
System.IO.File.WriteAllText(System.IO.Path.Combine(dir,"b.txt"),"b");
Console.WriteLine(System.IO.Directory.GetFiles(dir).Length);
System.IO.Directory.Delete(dir, true);
"#),
        &["2"]
    );
}

#[test]
fn get_temp_path_returns_non_empty_string() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.IO.Path.GetTempPath().Length > 0);"#),
        &["True"]
    );
}

#[test]
fn path_get_directory_name_returns_parent_path() {
    assert_eq!(
        run_csharp(r#"string dir = System.IO.Path.GetDirectoryName("/tmp/file.txt");
Console.WriteLine(dir);"#),
        &["/tmp"]
    );
}
