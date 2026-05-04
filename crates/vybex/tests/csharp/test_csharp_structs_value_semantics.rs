use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(default_struct_fields_start_at_zero, r#"struct Point { public int X; public int Y; } var point = new Point(); Console.WriteLine(point.X); Console.WriteLine(point.Y);"#, ["0", "0"]);
csharp_case!(struct_constructor_assigns_fields, r#"struct Point { public int X; public int Y; public Point(int x, int y) { X = x; Y = y; } } var point = new Point(2, 3); Console.WriteLine(point.X + point.Y);"#, ["5"]);
csharp_case!(struct_method_can_compute_from_fields, r#"struct Point { public int X; public int Y; public int Sum() { return X + Y; } } var point = new Point { X = 4, Y = 6 }; Console.WriteLine(point.Sum());"#, ["10"]);
csharp_case!(struct_property_can_return_derived_value, r#"struct Rect { public int W { get; set; } public int H { get; set; } public int Area => W * H; } var rect = new Rect { W = 3, H = 5 }; Console.WriteLine(rect.Area);"#, ["15"]);
csharp_case!(struct_assignment_copies_value_semantics, r#"struct Counter { public int Value; } var left = new Counter { Value = 1 }; var right = left; right.Value = 9; Console.WriteLine(left.Value); Console.WriteLine(right.Value);"#, ["1", "9"]);
csharp_case!(passing_struct_by_value_does_not_mutate_caller_copy, r#"struct Counter { public int Value; } void Bump(Counter counter) { counter.Value++; } var counter = new Counter { Value = 2 }; Bump(counter); Console.WriteLine(counter.Value);"#, ["2"]);
csharp_case!(passing_struct_by_ref_allows_mutation_of_original, r#"struct Counter { public int Value; } void Bump(ref Counter counter) { counter.Value++; } var counter = new Counter { Value = 2 }; Bump(ref counter); Console.WriteLine(counter.Value);"#, ["3"]);
csharp_case!(struct_can_implement_interface_contract, r#"interface IText { string Read(); } struct Token : IText { public string Read() { return "ok"; } } IText token = new Token(); Console.WriteLine(token.Read());"#, ["ok"]);
csharp_case!(struct_can_override_to_string_for_custom_output, r#"struct Token { public int Value; public override string ToString() { return "T:" + Value; } } Console.WriteLine(new Token { Value = 7 });"#, ["T:7"]);
csharp_case!(struct_equals_compares_field_values, r#"struct Point { public int X; public int Y; } var left = new Point { X = 1, Y = 2 }; var right = new Point { X = 1, Y = 2 }; Console.WriteLine(left.Equals(right));"#, ["True"]);
csharp_case!(nullable_struct_has_value_when_assigned, r#"System.DateTime? value = new System.DateTime(2024, 1, 1); Console.WriteLine(value.HasValue);"#, ["True"]);
csharp_case!(nullable_struct_value_exposes_underlying_member, r#"System.DateTime? value = new System.DateTime(2024, 1, 1); Console.WriteLine(value.Value.Year);"#, ["2024"]);
csharp_case!(struct_can_be_stored_inside_array, r#"struct Point { public int X; } var points = new[] { new Point { X = 3 }, new Point { X = 4 } }; foreach (var point in points) Console.WriteLine(point.X);"#, ["3", "4"]);
csharp_case!(struct_can_have_readonly_field_initialized_by_constructor, r#"struct Token { public readonly int Value; public Token(int value) { Value = value; } } Console.WriteLine(new Token(5).Value);"#, ["5"]);
csharp_case!(struct_can_contain_reference_type_field, r#"struct Wrapper { public string Name; } var wrapper = new Wrapper { Name = "text" }; Console.WriteLine(wrapper.Name);"#, ["text"]);
csharp_case!(struct_object_initializer_sets_auto_properties, r#"struct Box { public int Value { get; set; } } var box = new Box { Value = 11 }; Console.WriteLine(box.Value);"#, ["11"]);
csharp_case!(nested_struct_inside_class_is_constructible, r#"class Outer { public struct Inner { public int Value; } } var value = new Outer.Inner { Value = 8 }; Console.WriteLine(value.Value);"#, ["8"]);
csharp_case!(struct_can_have_static_member_shared_across_instances, r#"struct Token { public static int Count; public Token(int _) { Count++; } } new Token(1); new Token(2); Console.WriteLine(Token.Count);"#, ["2"]);
csharp_case!(readonly_struct_property_access_returns_value, r#"readonly struct Size { public int Width { get; } public Size(int width) { Width = width; } } Console.WriteLine(new Size(6).Width);"#, ["6"]);
csharp_case!(struct_can_implement_generic_interface, r#"interface IBox<T> { T Read(); } struct NumberBox : IBox<int> { public int Read() { return 14; } } IBox<int> box = new NumberBox(); Console.WriteLine(box.Read());"#, ["14"]);