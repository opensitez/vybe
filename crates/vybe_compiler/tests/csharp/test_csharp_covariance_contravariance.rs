//! Generic variance: `out` (covariant) and `in` (contravariant) type parameters.
use super::helpers::run_csharp;

#[test]
fn covariant_out_allows_derived_generic_argument_on_interface() {
    assert_eq!(
        run_csharp(
            r#"
interface IReader<out T> { T Read(); }
class StringReader : IReader<string> {
    public string Read() => "hello";
}
IReader<object> reader = new StringReader();
Console.WriteLine(reader.Read());
"#
        ),
        &["hello"]
    );
}

#[test]
fn contravariant_in_allows_base_generic_argument_on_interface() {
    assert_eq!(
        run_csharp(
            r#"
interface IWriter<in T> { void Write(T value); }
class ObjectWriter : IWriter<object> {
    public void Write(object value) => Console.WriteLine(value);
}
IWriter<string> writer = new ObjectWriter();
writer.Write("typed");
"#
        ),
        &["typed"]
    );
}

#[test]
fn array_covariance_allows_derived_array_in_base_array_reference() {
    assert_eq!(
        run_csharp(
            r#"
string[] strings = { "a", "b" };
object[] objects = strings;
Console.WriteLine(objects[0]);
"#
        ),
        &["a"]
    );
}

#[test]
fn ienumerable_covariance_allows_derived_sequence_in_base_reference() {
    assert_eq!(
        run_csharp(
            r#"
System.Collections.Generic.IEnumerable<string> strings =
    new System.Collections.Generic.List<string> { "x" };
System.Collections.Generic.IEnumerable<object> objects = strings;
foreach (var o in objects) Console.WriteLine(o);
"#
        ),
        &["x"]
    );
}

#[test]
fn func_return_type_covariance_allows_derived_func_in_base_func() {
    assert_eq!(
        run_csharp(
            r#"
System.Func<string> getString = () => "hi";
System.Func<object> getObject = getString;
Console.WriteLine(getObject());
"#
        ),
        &["hi"]
    );
}
