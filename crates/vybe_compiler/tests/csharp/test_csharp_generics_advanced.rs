//! Generic classes, methods, variance, default(T), and typeof with generics.
use super::helpers::run_csharp;

#[test]
fn generic_class_stores_and_returns_typed_value() {
    assert_eq!(
        run_csharp(
            r#"class Box<T> { public T Value; }
var b = new Box<int> { Value = 42 };
Console.WriteLine(b.Value);"#
        ),
        &["42"]
    );
}

#[test]
fn generic_method_infers_type_from_argument() {
    assert_eq!(
        run_csharp(
            r#"T Identity<T>(T value) => value;
Console.WriteLine(Identity(99));
Console.WriteLine(Identity("hi"));"#
        ),
        &["99", "hi"]
    );
}

#[test]
fn generic_stack_works_with_different_type_arguments() {
    assert_eq!(
        run_csharp(
            r#"class Stack<T> {
    System.Collections.Generic.List<T> items = new();
    public void Push(T v) => items.Add(v);
    public T Pop() { var v = items[items.Count-1]; items.RemoveAt(items.Count-1); return v; }
}
var s = new Stack<string>();
s.Push("a"); s.Push("b");
Console.WriteLine(s.Pop());"#
        ),
        &["b"]
    );
}

#[test]
fn default_of_generic_t_is_zero_for_value_types() {
    assert_eq!(
        run_csharp(
            r#"T Zero<T>() => default(T);
Console.WriteLine(Zero<int>());
Console.WriteLine(Zero<bool>());"#
        ),
        &["0", "False"]
    );
}

#[test]
fn default_of_generic_t_is_null_for_reference_types() {
    assert_eq!(
        run_csharp(
            r#"T Null<T>() where T : class => default(T);
Console.WriteLine(Null<string>() == null);"#
        ),
        &["True"]
    );
}

#[test]
fn typeof_on_closed_generic_includes_type_arg_in_name() {
    assert_eq!(
        run_csharp(
            r#"Console.WriteLine(typeof(System.Collections.Generic.List<int>).IsGenericType);"#
        ),
        &["True"]
    );
}

#[test]
fn generic_pair_swaps_values_through_method() {
    assert_eq!(
        run_csharp(
            r#"(T, T) Swap<T>(T a, T b) => (b, a);
var (x, y) = Swap(1, 2);
Console.WriteLine(x); Console.WriteLine(y);"#
        ),
        &["2", "1"]
    );
}

#[test]
fn generic_where_new_constraint_creates_instance_inside_method() {
    assert_eq!(
        run_csharp(
            r#"class Widget { public int Val = 5; }
T Make<T>() where T : new() => new T();
Console.WriteLine(Make<Widget>().Val);"#
        ),
        &["5"]
    );
}

#[test]
fn generic_list_works_with_interface_type_parameter() {
    assert_eq!(
        run_csharp(
            r#"interface IAnimal { string Sound(); }
class Cat : IAnimal { public string Sound() => "meow"; }
var animals = new System.Collections.Generic.List<IAnimal> { new Cat() };
foreach(var a in animals) Console.WriteLine(a.Sound());"#
        ),
        &["meow"]
    );
}
