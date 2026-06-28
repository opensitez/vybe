//! `System.Collections.Immutable`: ImmutableList, ImmutableDictionary, ImmutableArray.
use super::helpers::run_csharp;

#[test]
fn immutable_list_add_returns_new_list_old_unchanged() {
    assert_eq!(
        run_csharp(
            r#"var a=System.Collections.Immutable.ImmutableList<int>.Empty;
var b=a.Add(1).Add(2).Add(3);
Console.WriteLine(a.Count); Console.WriteLine(b.Count);"#
        ),
        &["0", "3"]
    );
}

#[test]
fn immutable_list_remove_returns_new_list() {
    assert_eq!(
        run_csharp(
            r#"var list=System.Collections.Immutable.ImmutableList.Create(1,2,3);
var smaller=list.Remove(2);
Console.WriteLine(list.Count); Console.WriteLine(smaller.Count);"#
        ),
        &["3", "2"]
    );
}

#[test]
fn immutable_array_indexer_reads_element() {
    assert_eq!(
        run_csharp(
            r#"var arr=System.Collections.Immutable.ImmutableArray.Create(10,20,30);
Console.WriteLine(arr[1]);"#
        ),
        &["20"]
    );
}

#[test]
fn immutable_dictionary_add_returns_new_dictionary() {
    assert_eq!(
        run_csharp(
            r#"var d=System.Collections.Immutable.ImmutableDictionary<string,int>.Empty;
var d2=d.Add("a",1).Add("b",2);
Console.WriteLine(d.Count); Console.WriteLine(d2["b"]);"#
        ),
        &["0", "2"]
    );
}

#[test]
fn immutable_list_set_item_replaces_at_index_returning_new() {
    assert_eq!(
        run_csharp(
            r#"var list=System.Collections.Immutable.ImmutableList.Create(1,2,3);
var updated=list.SetItem(1,99);
Console.WriteLine(list[1]); Console.WriteLine(updated[1]);"#
        ),
        &["2", "99"]
    );
}
