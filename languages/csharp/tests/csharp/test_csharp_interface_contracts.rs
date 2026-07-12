//! Standard interface contracts: IComparable, IEquatable, IEnumerable, IDisposable, ICloneable.
use super::helpers::run_csharp;

#[test]
fn icomparable_implementation_used_by_list_sort() {
    assert_eq!(
        run_csharp(
            r#"class Priority : System.IComparable<Priority> {
    public int Level;
    public int CompareTo(Priority other) => Level.CompareTo(other.Level);
}
var list = new System.Collections.Generic.List<Priority> {
    new Priority{Level=3}, new Priority{Level=1}, new Priority{Level=2}
};
list.Sort();
foreach(var p in list) Console.WriteLine(p.Level);"#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn iequatable_equals_compared_by_sorted_set_deduplication() {
    assert_eq!(
        run_csharp(
            r#"class Id : System.IEquatable<Id> {
    public int Value;
    public bool Equals(Id other) => other?.Value == Value;
    public override bool Equals(object o) => o is Id i && Equals(i);
    public override int GetHashCode() => Value;
}
var set = new System.Collections.Generic.HashSet<Id>(
    System.Collections.Generic.EqualityComparer<Id>.Default);
set.Add(new Id{Value=1}); set.Add(new Id{Value=1});
Console.WriteLine(set.Count);"#
        ),
        &["1"]
    );
}

#[test]
fn idisposable_dispose_called_by_using_statement() {
    assert_eq!(
        run_csharp(
            r#"class Resource : System.IDisposable {
    public static int Disposed = 0;
    public void Dispose() => Disposed++;
}
using(var r = new Resource()) { }
Console.WriteLine(Resource.Disposed);"#
        ),
        &["1"]
    );
}

#[test]
fn ienumerable_foreach_drives_custom_iterator() {
    assert_eq!(
        run_csharp(
            r#"class Counter : System.Collections.Generic.IEnumerable<int> {
    public System.Collections.Generic.IEnumerator<int> GetEnumerator() {
        yield return 1; yield return 2; yield return 3;
    }
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}
int sum=0;
foreach(var n in new Counter()) sum+=n;
Console.WriteLine(sum);"#
        ),
        &["6"]
    );
}

#[test]
fn icloneable_clone_returns_independent_copy() {
    assert_eq!(
        run_csharp(
            r#"class Box : System.ICloneable {
    public int Value;
    public object Clone() => new Box { Value = Value };
}
var original = new Box { Value=5 };
var copy = (Box)original.Clone();
copy.Value = 99;
Console.WriteLine(original.Value);"#
        ),
        &["5"]
    );
}

#[test]
fn iformattable_tostring_with_format_spec() {
    assert_eq!(
        run_csharp(
            r#"System.IFormattable value = (object)3.14;
Console.WriteLine(value.ToString("F1", System.Globalization.CultureInfo.InvariantCulture));"#
        ),
        &["3.1"]
    );
}
