//! Generic method type inference, constraints at call sites, and open vs
//! closed constructed types.
use super::helpers::run_csharp;

#[test]
fn generic_method_infers_type_argument_from_parameter() {
    assert_eq!(
        run_csharp(
            r#"
T Identity<T>(T value) { return value; }
Console.WriteLine(Identity(42));
Console.WriteLine(Identity("text"));
"#
        ),
        &["42", "text"]
    );
}

#[test]
fn generic_method_infers_type_from_return_assignment() {
    assert_eq!(
        run_csharp(
            r#"
T First<T>(T left, T right) { return left; }
string chosen = First("left", "right");
Console.WriteLine(chosen);
"#
        ),
        &["left"]
    );
}

#[test]
fn generic_class_method_infers_from_arguments() {
    assert_eq!(
        run_csharp(
            r#"
class Box<T> {
    public T Value;
    public Box(T value) { Value = value; }
    public T Get() { return Value; }
}
var numbers = new Box<int>(5);
Console.WriteLine(numbers.Get());
"#
        ),
        &["5"]
    );
}

#[test]
fn generic_method_with_multiple_type_params_infers_both() {
    assert_eq!(
        run_csharp(
            r#"
(K, V) MakePair<K, V>(K key, V value) { return (key, value); }
var pair = MakePair("id", 7);
Console.WriteLine(pair.Item1);
Console.WriteLine(pair.Item2);
"#
        ),
        &["id", "7"]
    );
}

#[test]
fn generic_collection_preserves_element_type_through_add_and_index() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
var scores = new Dictionary<string, int>();
scores.Add("ada", 99);
scores.Add("lin", 88);
Console.WriteLine(scores["ada"]);
Console.WriteLine(scores.ContainsKey("lin"));
"#
        ),
        &["99", "True"]
    );
}

#[test]
fn generic_method_constraint_class_allows_reference_type_members() {
    assert_eq!(
        run_csharp(
            r#"
string Describe<T>(T value) where T : class {
    return value == null ? "null" : value.ToString();
}
Console.WriteLine(Describe("data"));
"#
        ),
        &["data"]
    );
}

#[test]
fn generic_method_constraint_struct_accepts_unboxed_value_type() {
    assert_eq!(
        run_csharp(
            r#"
int Scale<T>(T value) where T : struct {
    return 2 * (int)(object)value;
}
Console.WriteLine(Scale(6));
"#
        ),
        &["12"]
    );
}

#[test]
fn generic_method_new_constraint_allows_parameterless_construction() {
    assert_eq!(
        run_csharp(
            r#"
class Widget { public int Size = 4; }
T Create<T>() where T : new() { return new T(); }
Console.WriteLine(Create<Widget>().Size);
"#
        ),
        &["4"]
    );
}

#[test]
fn generic_interface_implementation_is_visible_through_type_parameter() {
    assert_eq!(
        run_csharp(
            r#"
interface IReader {
    int Read();
}
class Sensor : IReader {
    public int Read() { return 17; }
}
int Load<T>(T device) where T : IReader { return device.Read(); }
Console.WriteLine(Load(new Sensor()));
"#
        ),
        &["17"]
    );
}

#[test]
fn generic_nested_type_shares_outer_type_argument() {
    assert_eq!(
        run_csharp(
            r#"
class Outer<T> {
    public class Inner {
        public T Value;
    }
    public Inner Build(T value) {
        return new Inner { Value = value };
    }
}
var built = new Outer<string>().Build("nested");
Console.WriteLine(built.Value);
"#
        ),
        &["nested"]
    );
}

#[test]
fn generic_method_overload_resolution_prefers_specific_argument_types() {
    assert_eq!(
        run_csharp(
            r#"
string Pick(int value) { return "int:" + value; }
string Pick(string value) { return "str:" + value; }
Console.WriteLine(Pick(3));
Console.WriteLine(Pick("3"));
"#
        ),
        &["int:3", "str:3"]
    );
}

#[test]
fn covariant_array_assignment_allows_derived_elements_in_object_array() {
    assert_eq!(
        run_csharp(
            r#"
class Fruit { public string Name; }
class Apple : Fruit { }
Fruit[] basket = new Apple[2];
basket[0] = new Apple { Name = "fuji" };
Console.WriteLine(basket[0].Name);
"#
        ),
        &["fuji"]
    );
}
