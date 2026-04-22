use super::helpers::run_csharp;

#[test]
fn fully_qualified_thread_sleep_uses_shared_dotnet_surface() {
    let out = run_csharp(r#"
        Console.WriteLine("before");
        System.Threading.Thread.Sleep(1);
        Console.WriteLine("after");
    "#);
    assert_eq!(out, vec!["before", "after"]);
}