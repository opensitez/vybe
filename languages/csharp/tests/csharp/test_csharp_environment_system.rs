//! `System.Environment`, `System.Console`, `System.GC`, `System.AppDomain`.
use super::helpers::run_csharp;

#[test]
fn environment_newline_is_non_empty() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Environment.NewLine.Length>0);"#),
        &["True"]
    );
}

#[test]
fn environment_processor_count_is_positive() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Environment.ProcessorCount>0);"#),
        &["True"]
    );
}

#[test]
fn environment_get_environment_variable_returns_null_for_unknown() {
    assert_eq!(
        run_csharp(
            r#"var v=System.Environment.GetEnvironmentVariable("__VYBE_NOSUCH_VAR__123");
Console.WriteLine(v==null);"#
        ),
        &["True"]
    );
}

#[test]
fn gc_collect_runs_without_error() {
    assert_eq!(
        run_csharp(
            r#"System.GC.Collect();
Console.WriteLine("ok");"#
        ),
        &["ok"]
    );
}

#[test]
fn gc_get_total_memory_returns_positive_long() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.GC.GetTotalMemory(false)>0);"#),
        &["True"]
    );
}

#[test]
fn appdomain_current_domain_not_null() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.AppDomain.CurrentDomain!=null);"#),
        &["True"]
    );
}
