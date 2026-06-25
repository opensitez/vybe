//! `ConcurrentDictionary<TKey,TValue>` thread-safe operations.
use super::helpers::run_csharp;

#[test]
fn try_add_inserts_when_key_absent() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Concurrent.ConcurrentDictionary<string,int>();
Console.WriteLine(d.TryAdd("a", 1));
Console.WriteLine(d["a"]);"#
        ),
        &["True", "1"]
    );
}

#[test]
fn try_add_returns_false_when_key_present() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Concurrent.ConcurrentDictionary<string,int>();
d.TryAdd("a", 1);
Console.WriteLine(d.TryAdd("a", 9));"#
        ),
        &["False"]
    );
}

#[test]
fn get_or_add_returns_existing_value_without_adding() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Concurrent.ConcurrentDictionary<string,int>();
d["x"] = 5;
Console.WriteLine(d.GetOrAdd("x", 99));"#
        ),
        &["5"]
    );
}

#[test]
fn get_or_add_inserts_when_key_absent() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Concurrent.ConcurrentDictionary<string,int>();
Console.WriteLine(d.GetOrAdd("new", 42));"#
        ),
        &["42"]
    );
}

#[test]
fn add_or_update_replaces_existing_via_factory() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Concurrent.ConcurrentDictionary<string,int>();
d["k"] = 1;
d.AddOrUpdate("k", 0, (key, old) => old + 10);
Console.WriteLine(d["k"]);"#
        ),
        &["11"]
    );
}

#[test]
fn try_remove_extracts_and_deletes_entry() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Concurrent.ConcurrentDictionary<string,int>();
d["x"] = 7;
Console.WriteLine(d.TryRemove("x", out int v));
Console.WriteLine(v);
Console.WriteLine(d.Count);"#
        ),
        &["True", "7", "0"]
    );
}
