use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(
    generic_method_with_class_constraint_accepts_reference_type,
    r#"string Echo<T>(T value) where T : class { return value.ToString(); } Console.WriteLine(Echo("text"));"#,
    ["text"]
);
csharp_case!(
    generic_method_with_struct_constraint_accepts_value_type,
    r#"string Wrap<T>(T value) where T : struct { return value.ToString(); } Console.WriteLine(Wrap(5));"#,
    ["5"]
);
csharp_case!(
    generic_method_with_new_constraint_constructs_instance,
    r#"class Box { public int Value = 9; } T Create<T>() where T : new() { return new T(); } Console.WriteLine(Create<Box>().Value);"#,
    ["9"]
);
csharp_case!(
    generic_method_with_interface_constraint_calls_member,
    r#"interface ILabel { string Label(); } class Item : ILabel { public string Label() { return "ok"; } } string Read<T>(T value) where T : ILabel { return value.Label(); } Console.WriteLine(Read(new Item()));"#,
    ["ok"]
);
csharp_case!(
    generic_method_with_base_class_constraint_accesses_base_member,
    r#"class Base { public string Name = "base"; } class Child : Base { } string Read<T>(T value) where T : Base { return value.Name; } Console.WriteLine(Read(new Child()));"#,
    ["base"]
);
csharp_case!(
    generic_method_with_multiple_constraints_uses_interface_and_constructor,
    r#"interface IValue { int Read(); } class Item : IValue { public int Read() { return 4; } } int Build<T>() where T : IValue, new() { return new T().Read(); } Console.WriteLine(Build<Item>());"#,
    ["4"]
);
csharp_case!(
    generic_class_with_constraint_can_store_value,
    r#"class Holder<T> where T : class { public T Value { get; set; } } var holder = new Holder<string> { Value = "abc" }; Console.WriteLine(holder.Value);"#,
    ["abc"]
);
csharp_case!(
    generic_class_with_new_constraint_can_create_member,
    r#"class Factory<T> where T : new() { public T Build() { return new T(); } } class Item { public string Name = "built"; } Console.WriteLine(new Factory<Item>().Build().Name);"#,
    ["built"]
);
csharp_case!(
    generic_method_returns_default_for_value_type,
    r#"T Zero<T>() where T : struct { return default(T); } Console.WriteLine(Zero<int>());"#,
    ["0"]
);
csharp_case!(
    generic_method_returns_default_for_reference_type,
    r#"T Empty<T>() where T : class { return default(T); } Console.WriteLine(Empty<string>() is null);"#,
    ["True"]
);
csharp_case!(
    generic_interface_implementation_preserves_type_argument,
    r#"interface IBox<T> { T Read(); } class NumberBox : IBox<int> { public int Read() { return 8; } } Console.WriteLine(((IBox<int>)new NumberBox()).Read());"#,
    ["8"]
);
csharp_case!(
    generic_method_can_compare_two_values_with_equality,
    r#"bool Same<T>(T left, T right) { return left.Equals(right); } Console.WriteLine(Same(3, 3));"#,
    ["True"]
);
csharp_case!(
    generic_method_with_where_t_base_class_works_on_derived_input,
    r#"class Person { public string Name = "Ada"; } class Admin : Person { } string Read<T>(T person) where T : Person { return person.Name; } Console.WriteLine(Read(new Admin()));"#,
    ["Ada"]
);
csharp_case!(
    generic_method_can_use_list_of_t_argument,
    r#"using System.Collections.Generic; int Count<T>(List<T> items) { return items.Count; } Console.WriteLine(Count(new List<string> { "a", "b" }));"#,
    ["2"]
);
csharp_case!(
    generic_static_field_is_independent_per_closed_type,
    r#"class Counter<T> { public static int Value; } Counter<int>.Value = 2; Counter<string>.Value = 5; Console.WriteLine(Counter<int>.Value); Console.WriteLine(Counter<string>.Value);"#,
    ["2", "5"]
);
csharp_case!(
    generic_method_with_constraint_can_read_property_from_interface,
    r#"interface INamed { string Name { get; } } class User : INamed { public string Name => "Grace"; } string Read<T>(T item) where T : INamed { return item.Name; } Console.WriteLine(Read(new User()));"#,
    ["Grace"]
);
csharp_case!(
    generic_method_can_return_tuple_of_type_argument_and_count,
    r#"(T, int) Pair<T>(T value) { return (value, 1); } var result = Pair("x"); Console.WriteLine(result.Item1 + result.Item2);"#,
    ["x1"]
);
csharp_case!(
    generic_method_can_swap_two_values_by_tuple_assignment,
    r#"(T, T) Swap<T>(T left, T right) { (left, right) = (right, left); return (left, right); } var result = Swap(1, 9); Console.WriteLine(result.Item1); Console.WriteLine(result.Item2);"#,
    ["9", "1"]
);
csharp_case!(
    generic_method_with_struct_constraint_can_add_nullable_check,
    r#"string Describe<T>(T? value) where T : struct { return value.HasValue ? value.Value.ToString() : "none"; } Console.WriteLine(Describe<int>(7));"#,
    ["7"]
);
csharp_case!(
    generic_class_with_base_constraint_can_call_virtual_method,
    r#"class Base { public virtual string Read() { return "base"; } } class Child : Base { public override string Read() { return "child"; } } class Reader<T> where T : Base { public string Run(T value) { return value.Read(); } } Console.WriteLine(new Reader<Child>().Run(new Child()));"#,
    ["child"]
);
