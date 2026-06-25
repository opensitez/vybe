//! Generic collection behaviour: capacity, trimming, EnsureCapacity, Contains semantics.
use super::helpers::run_csharp;

#[test]
fn list_capacity_doubles_on_overflow() {
    assert_eq!(
        run_csharp(r#"var list=new System.Collections.Generic.List<int>(4);
for(int i=0;i<8;i++) list.Add(i);
Console.WriteLine(list.Count); Console.WriteLine(list.Capacity>=8);"#),
        &["8", "True"]
    );
}

#[test]
fn list_trim_excess_reduces_capacity_to_count() {
    assert_eq!(
        run_csharp(r#"var list=new System.Collections.Generic.List<int>(100);
list.Add(1); list.Add(2);
list.TrimExcess();
Console.WriteLine(list.Capacity<=list.Count*1);"#),
        &["True"]
    );
}

#[test]
fn dictionary_contains_key_vs_contains_value() {
    assert_eq!(
        run_csharp(r#"var d=new System.Collections.Generic.Dictionary<string,int>{{"a",1}};
Console.WriteLine(d.ContainsKey("a"));
Console.WriteLine(d.ContainsValue(1));
Console.WriteLine(d.ContainsKey("b"));"#),
        &["True", "True", "False"]
    );
}

#[test]
fn hash_set_union_with_intersects_and_except() {
    assert_eq!(
        run_csharp(r#"var a=new System.Collections.Generic.HashSet<int>{1,2,3,4};
var b=new System.Collections.Generic.HashSet<int>{3,4,5,6};
a.IntersectWith(b);
Console.WriteLine(a.Count); Console.WriteLine(a.Contains(3));"#),
        &["2", "True"]
    );
}

#[test]
fn sorted_set_get_view_between_returns_subset() {
    assert_eq!(
        run_csharp(r#"var s=new System.Collections.Generic.SortedSet<int>{1,2,3,4,5};
var view=s.GetViewBetween(2,4);
Console.WriteLine(view.Count);"#),
        &["3"]
    );
}
