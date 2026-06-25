//! Custom `IComparable<T>` and `IComparer<T>` for sorting and ordering.
use super::helpers::run_csharp;

#[test]
fn custom_icomparable_used_by_array_sort() {
    assert_eq!(
        run_csharp(
            r#"class Score : System.IComparable<Score> {
    public int Value;
    public int CompareTo(Score other) => Value.CompareTo(other.Value);
}
var scores = new[]{
    new Score{Value=5}, new Score{Value=1}, new Score{Value=3}
};
System.Array.Sort(scores);
Console.WriteLine(scores[0].Value);
Console.WriteLine(scores[2].Value);"#
        ),
        &["1", "5"]
    );
}

#[test]
fn comparer_default_sorts_strings_lexicographically() {
    assert_eq!(
        run_csharp(
            r#"var list = new System.Collections.Generic.List<string>{"banana","apple","cherry"};
list.Sort(System.StringComparer.Ordinal);
Console.WriteLine(list[0]);"#
        ),
        &["apple"]
    );
}

#[test]
fn custom_icomparer_reverses_natural_order() {
    assert_eq!(
        run_csharp(
            r#"class Desc : System.Collections.Generic.IComparer<int> {
    public int Compare(int x, int y) => y.CompareTo(x);
}
var list = new System.Collections.Generic.List<int>{3,1,4,1,5};
list.Sort(new Desc());
Console.WriteLine(list[0]);"#
        ),
        &["5"]
    );
}

#[test]
fn linq_order_by_uses_default_icomparable_for_value_types() {
    assert_eq!(
        run_csharp(
            r#"var result = new[]{3,1,2}.OrderBy(x=>x);
foreach(var n in result) Console.WriteLine(n);"#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn comparer_create_builds_comparer_from_lambda() {
    assert_eq!(
        run_csharp(
            r#"var cmp = System.Collections.Generic.Comparer<string>.Create(
    (a,b) => a.Length.CompareTo(b.Length));
var list = new System.Collections.Generic.List<string>{"cc","aaa","b"};
list.Sort(cmp);
Console.WriteLine(list[0]);"#
        ),
        &["b"]
    );
}
