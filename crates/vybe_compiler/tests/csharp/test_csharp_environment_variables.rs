//! `Environment.SetEnvironmentVariable` / `GetEnvironmentVariable` round-trips.
use super::helpers::run_csharp;

#[test]
fn set_and_get_environment_variable_roundtrip() {
    assert_eq!(
        run_csharp(r#"System.Environment.SetEnvironmentVariable("VYBE_TEST_KEY","hello");
Console.WriteLine(System.Environment.GetEnvironmentVariable("VYBE_TEST_KEY"));"#),
        &["hello"]
    );
}

#[test]
fn deleting_environment_variable_makes_it_null() {
    assert_eq!(
        run_csharp(r#"System.Environment.SetEnvironmentVariable("VYBE_DEL_KEY","x");
System.Environment.SetEnvironmentVariable("VYBE_DEL_KEY",null);
Console.WriteLine(System.Environment.GetEnvironmentVariable("VYBE_DEL_KEY")==null);"#),
        &["True"]
    );
}

#[test]
fn current_directory_is_non_empty_string() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Environment.CurrentDirectory.Length>0);"#),
        &["True"]
    );
}
