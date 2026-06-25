//! `Dictionary<TKey,TValue>` full API coverage.
use super::helpers::run_csharp;

#[test]
fn add_inserts_key_value_pair_and_count_increases() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Generic.Dictionary<string,int>();
d.Add("a",1); d.Add("b",2);
Console.WriteLine(d.Count);"#
        ),
        &["2"]
    );
}

#[test]
fn indexer_set_replaces_value_for_existing_key() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Generic.Dictionary<string,int>();
d["x"] = 1; d["x"] = 9;
Console.WriteLine(d["x"]);"#
        ),
        &["9"]
    );
}

#[test]
fn contains_key_returns_false_for_absent_key() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Generic.Dictionary<string,int>{{"a",1}};
Console.WriteLine(d.ContainsKey("z"));"#
        ),
        &["False"]
    );
}

#[test]
fn contains_value_finds_value_regardless_of_key() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Generic.Dictionary<string,int>{{"a",42}};
Console.WriteLine(d.ContainsValue(42));"#
        ),
        &["True"]
    );
}

#[test]
fn try_get_value_returns_true_and_out_value_on_hit() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Generic.Dictionary<string,int>{{"k",5}};
Console.WriteLine(d.TryGetValue("k", out int v));
Console.WriteLine(v);"#
        ),
        &["True", "5"]
    );
}

#[test]
fn try_get_value_returns_false_and_default_on_miss() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Generic.Dictionary<string,int>();
Console.WriteLine(d.TryGetValue("nope", out int v));
Console.WriteLine(v);"#
        ),
        &["False", "0"]
    );
}

#[test]
fn remove_deletes_key_and_reduces_count() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Generic.Dictionary<string,int>{{"a",1},{"b",2}};
d.Remove("a");
Console.WriteLine(d.Count);"#
        ),
        &["1"]
    );
}

#[test]
fn keys_collection_enumerates_all_inserted_keys() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Generic.Dictionary<string,int>{{"x",1},{"y",2}};
var keys = new System.Collections.Generic.List<string>(d.Keys);
keys.Sort();
foreach(var k in keys) Console.WriteLine(k);"#
        ),
        &["x", "y"]
    );
}

#[test]
fn values_collection_sum_matches_expected_total() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Generic.Dictionary<string,int>{{"a",3},{"b",7}};
int sum=0; foreach(var v in d.Values) sum+=v;
Console.WriteLine(sum);"#
        ),
        &["10"]
    );
}

#[test]
fn foreach_over_dictionary_yields_key_value_pairs() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Generic.Dictionary<int,int>{{1,10}};
foreach(var pair in d) Console.WriteLine(pair.Key + ":" + pair.Value);"#
        ),
        &["1:10"]
    );
}

#[test]
fn clear_removes_all_entries() {
    assert_eq!(
        run_csharp(
            r#"var d = new System.Collections.Generic.Dictionary<string,int>{{"a",1}};
d.Clear();
Console.WriteLine(d.Count);"#
        ),
        &["0"]
    );
}
