use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(is_pattern_captures_string_value_for_length_check, r#"object item = "alpha"; if (item is string text) Console.WriteLine(text.Length);"#, ["5"]);
csharp_case!(is_pattern_rejects_non_matching_type, r#"object item = 42; Console.WriteLine(item is string text);"#, ["False"]);
csharp_case!(is_not_null_pattern_filters_null_reference, r#"string text = null; Console.WriteLine(text is not null);"#, ["False"]);
csharp_case!(is_not_null_pattern_accepts_non_null_reference, r#"string text = "ready"; Console.WriteLine(text is not null);"#, ["True"]);
csharp_case!(switch_statement_type_pattern_matches_string_arm, r#"object item = "beta"; switch (item) { case string text: Console.WriteLine(text.ToUpper()); break; default: Console.WriteLine("other"); break; }"#, ["BETA"]);
csharp_case!(switch_statement_type_pattern_matches_int_arm, r#"object item = 9; switch (item) { case string text: Console.WriteLine(text); break; case int number: Console.WriteLine(number * 3); break; default: Console.WriteLine("other"); break; }"#, ["27"]);
csharp_case!(switch_statement_when_guard_matches_large_number, r#"var x = 12; switch (x) { case int number when number > 10: Console.WriteLine("large"); break; case int number: Console.WriteLine("small"); break; }"#, ["large"]);
csharp_case!(switch_statement_when_guard_matches_small_number, r#"var x = 3; switch (x) { case int number when number > 10: Console.WriteLine("large"); break; case int number: Console.WriteLine("small"); break; }"#, ["small"]);
csharp_case!(constant_pattern_matches_true_boolean, r#"object value = true; if (value is true) Console.WriteLine("yes");"#, ["yes"]);
csharp_case!(constant_pattern_matches_false_boolean, r#"object value = false; if (value is false) Console.WriteLine("no");"#, ["no"]);
csharp_case!(var_pattern_binds_any_value_in_switch_statement, r#"object item = 18; switch (item) { case var anything: Console.WriteLine(anything); break; }"#, ["18"]);
csharp_case!(positional_tuple_pattern_matches_exact_pair, r#"var pair = (2, 3); if (pair is (2, 3)) Console.WriteLine("match");"#, ["match"]);
csharp_case!(positional_tuple_pattern_with_discard_matches_partial_pair, r#"var pair = (2, 9); if (pair is (2, _)) Console.WriteLine("left-two");"#, ["left-two"]);
csharp_case!(relational_pattern_matches_value_in_if_statement, r#"var score = 91; if (score is >= 90) Console.WriteLine("A");"#, ["A"]);
csharp_case!(relational_pattern_with_range_match_in_if_statement, r#"var score = 85; if (score is >= 80 and < 90) Console.WriteLine("B");"#, ["B"]);
csharp_case!(object_type_pattern_matches_base_class_instance, r#"class Animal { } class Dog : Animal { } object pet = new Dog(); Console.WriteLine(pet is Animal);"#, ["True"]);
csharp_case!(declaration_pattern_on_nullable_value_extracts_number, r#"int? value = 7; if (value is int number) Console.WriteLine(number + 1);"#, ["8"]);
csharp_case!(null_pattern_matches_missing_nullable_value, r#"int? value = null; if (value is null) Console.WriteLine("missing");"#, ["missing"]);
csharp_case!(type_pattern_in_ternary_expression_selects_branch, r#"object item = "cs"; Console.WriteLine(item is string ? "text" : "other");"#, ["text"]);
csharp_case!(property_pattern_like_access_via_guarded_type_check, r#"class Point { public int X { get; set; } public int Y { get; set; } } object item = new Point { X = 5, Y = 8 }; if (item is Point point && point.X == 5) Console.WriteLine(point.Y);"#, ["8"]);