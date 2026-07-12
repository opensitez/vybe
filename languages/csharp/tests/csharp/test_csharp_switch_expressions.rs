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
    switch_expression_matches_small_integer_constant,
    r#"var x = 2; Console.WriteLine(x switch { 1 => "one", 2 => "two", _ => "other" });"#,
    ["two"]
);
csharp_case!(
    switch_expression_uses_default_discard_arm_for_unknown_value,
    r#"var x = 9; Console.WriteLine(x switch { 1 => "one", 2 => "two", _ => "other" });"#,
    ["other"]
);
csharp_case!(
    switch_expression_handles_negative_number_with_relational_arm,
    r#"var x = -3; Console.WriteLine(x switch { < 0 => "neg", 0 => "zero", > 0 => "pos" });"#,
    ["neg"]
);
csharp_case!(
    switch_expression_distinguishes_zero_from_positive_values,
    r#"var x = 0; Console.WriteLine(x switch { < 0 => "neg", 0 => "zero", > 0 => "pos" });"#,
    ["zero"]
);
csharp_case!(
    switch_expression_matches_string_literal_cases,
    r#"var word = "beta"; Console.WriteLine(word switch { "alpha" => "A", "beta" => "B", _ => "?" });"#,
    ["B"]
);
csharp_case!(
    switch_expression_matches_char_literal_cases,
    r#"char ch = 'x'; Console.WriteLine(ch switch { 'x' => "ex", 'y' => "why", _ => "other" });"#,
    ["ex"]
);
csharp_case!(
    switch_expression_orders_tuple_pattern_arms,
    r#"var pair = (1, 0); Console.WriteLine(pair switch { (0, 0) => "origin", (1, 0) => "unit-x", _ => "other" });"#,
    ["unit-x"]
);
csharp_case!(
    switch_expression_matches_boolean_tuple_combinations,
    r#"var flags = (true, false); Console.WriteLine(flags switch { (true, true) => "both", (true, false) => "left", _ => "other" });"#,
    ["left"]
);
csharp_case!(
    switch_expression_returns_interpolated_string_from_arm,
    r#"var score = 87; Console.WriteLine(score switch { >= 90 => $"A:{score}", >= 80 => $"B:{score}", _ => $"C:{score}" });"#,
    ["B:87"]
);
csharp_case!(
    switch_expression_uses_parenthesized_result_expression,
    r#"var x = 4; Console.WriteLine(x switch { 4 => (2 + 3), _ => 0 });"#,
    ["5"]
);
csharp_case!(
    switch_expression_matches_nullable_with_null_arm,
    r#"int? value = null; Console.WriteLine(value switch { null => "missing", 0 => "zero", _ => "number" });"#,
    ["missing"]
);
csharp_case!(
    switch_expression_matches_nullable_with_value_arm,
    r#"int? value = 12; Console.WriteLine(value switch { null => "missing", > 10 => "large", _ => "small" });"#,
    ["large"]
);
csharp_case!(
    switch_expression_handles_enum_like_constants,
    r#"enum State { Idle, Running, Done } var state = State.Done; Console.WriteLine(state switch { State.Idle => "idle", State.Running => "running", State.Done => "done", _ => "other" });"#,
    ["done"]
);
csharp_case!(
    switch_expression_matches_object_type_pattern,
    r#"object item = "hello"; Console.WriteLine(item switch { string text => text.ToUpper(), int number => (number * 2).ToString(), _ => "other" });"#,
    ["HELLO"]
);
csharp_case!(
    switch_expression_matches_integer_type_pattern,
    r#"object item = 7; Console.WriteLine(item switch { string text => text, int number => (number + 1).ToString(), _ => "other" });"#,
    ["8"]
);
csharp_case!(
    switch_expression_uses_when_guard_for_even_value,
    r#"var x = 8; Console.WriteLine(x switch { int n when n % 2 == 0 => "even", int n => "odd" });"#,
    ["even"]
);
csharp_case!(
    switch_expression_uses_when_guard_for_odd_value,
    r#"var x = 5; Console.WriteLine(x switch { int n when n % 2 == 0 => "even", int n => "odd" });"#,
    ["odd"]
);
csharp_case!(
    switch_expression_combines_length_check_and_content,
    r#"var text = "tool"; Console.WriteLine(text switch { string s when s.Length == 4 => "len4", string s => s, _ => "none" });"#,
    ["len4"]
);
csharp_case!(
    switch_expression_handles_tuple_with_discard_component,
    r#"var pair = (3, 9); Console.WriteLine(pair switch { (3, _) => "starts-three", (_, 9) => "ends-nine", _ => "other" });"#,
    ["starts-three"]
);
csharp_case!(
    switch_expression_falls_through_to_second_tuple_arm,
    r#"var pair = (4, 9); Console.WriteLine(pair switch { (3, _) => "starts-three", (_, 9) => "ends-nine", _ => "other" });"#,
    ["ends-nine"]
);
