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
    implicit_numeric_conversion_from_int_to_double,
    r#"int count = 7; double total = count; Console.WriteLine(total);"#,
    ["7"]
);
csharp_case!(
    explicit_numeric_conversion_from_double_to_int,
    r#"double value = 7.9; int whole = (int)value; Console.WriteLine(whole);"#,
    ["7"]
);
csharp_case!(
    boxing_value_type_into_object_and_printing_runtime_value,
    r#"int count = 11; object boxed = count; Console.WriteLine(boxed);"#,
    ["11"]
);
csharp_case!(
    unboxing_object_back_to_integer_value,
    r#"object boxed = 21; int count = (int)boxed; Console.WriteLine(count + 1);"#,
    ["22"]
);
csharp_case!(
    as_operator_returns_string_instance_for_matching_type,
    r#"object item = "hello"; string text = item as string; Console.WriteLine(text);"#,
    ["hello"]
);
csharp_case!(
    as_operator_returns_null_for_non_matching_type,
    r#"object item = 42; string text = item as string; Console.WriteLine(text is null);"#,
    ["True"]
);
csharp_case!(
    is_operator_reports_true_for_assignable_interface,
    r#"using System.Collections.Generic; object item = new List<int>(); Console.WriteLine(item is IEnumerable<int>);"#,
    ["True"]
);
csharp_case!(
    is_operator_reports_false_for_unrelated_reference_type,
    r#"object item = "text"; Console.WriteLine(item is System.DateTime);"#,
    ["False"]
);
csharp_case!(
    convert_class_parses_integer_from_string,
    r#"Console.WriteLine(System.Convert.ToInt32("42") + 8);"#,
    ["50"]
);
csharp_case!(
    convert_class_creates_double_from_integer,
    r#"Console.WriteLine(System.Convert.ToDouble(5));"#,
    ["5"]
);
csharp_case!(
    char_to_integer_cast_produces_code_point,
    r#"char ch = 'A'; Console.WriteLine((int)ch);"#,
    ["65"]
);
csharp_case!(
    integer_to_char_cast_produces_character,
    r#"int value = 66; Console.WriteLine((char)value);"#,
    ["B"]
);
csharp_case!(
    boxing_nullable_with_value_prints_number,
    r#"int? value = 13; object boxed = value; Console.WriteLine(boxed);"#,
    ["13"]
);
csharp_case!(
    casting_object_to_interface_allows_method_dispatch,
    r#"interface IGreeter { string Say(); } class Greeter : IGreeter { public string Say() { return "hi"; } } object item = new Greeter(); Console.WriteLine(((IGreeter)item).Say());"#,
    ["hi"]
);
csharp_case!(
    casting_object_to_base_class_exposes_virtual_member,
    r#"class Base { public virtual string Name() { return "base"; } } class Child : Base { public override string Name() { return "child"; } } object item = new Child(); Console.WriteLine(((Base)item).Name());"#,
    ["child"]
);
csharp_case!(
    nullable_value_type_is_pattern_extracts_underlying_number,
    r#"int? maybe = 30; if (maybe is int value) Console.WriteLine(value / 3);"#,
    ["10"]
);
csharp_case!(
    string_interpolation_converts_integer_to_text,
    r#"int count = 9; Console.WriteLine($"count={count}");"#,
    ["count=9"]
);
csharp_case!(
    enum_can_be_cast_to_underlying_integer,
    r#"enum Mode { Off = 0, On = 5 } Console.WriteLine((int)Mode.On);"#,
    ["5"]
);
csharp_case!(
    underlying_integer_can_be_cast_back_to_enum,
    r#"enum Mode { Off = 0, On = 5 } var mode = (Mode)5; Console.WriteLine(mode);"#,
    ["On"]
);
csharp_case!(
    reference_conversion_from_derived_to_base_keeps_overrides,
    r#"class Animal { public virtual string Speak() { return "animal"; } } class Dog : Animal { public override string Speak() { return "woof"; } } Dog dog = new Dog(); Animal animal = dog; Console.WriteLine(animal.Speak());"#,
    ["woof"]
);
