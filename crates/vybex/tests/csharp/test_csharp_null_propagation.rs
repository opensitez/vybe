use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(null_conditional_property_access_returns_value_for_non_null_object, r#"class User { public string Name { get; set; } } var user = new User { Name = "Ada" }; Console.WriteLine(user?.Name);"#, ["Ada"]);
csharp_case!(null_conditional_property_access_returns_null_for_missing_object, r#"class User { public string Name { get; set; } } User user = null; Console.WriteLine(user?.Name ?? "none");"#, ["none"]);
csharp_case!(null_conditional_method_call_returns_value_for_non_null_object, r#"string text = "hello"; Console.WriteLine(text?.ToUpper());"#, ["HELLO"]);
csharp_case!(null_conditional_method_call_returns_fallback_for_null_object, r#"string text = null; Console.WriteLine(text?.ToUpper() ?? "empty");"#, ["empty"]);
csharp_case!(null_coalescing_operator_selects_left_when_non_null, r#"string value = "left"; Console.WriteLine(value ?? "right");"#, ["left"]);
csharp_case!(null_coalescing_operator_selects_right_when_left_is_null, r#"string value = null; Console.WriteLine(value ?? "right");"#, ["right"]);
csharp_case!(null_coalescing_assignment_sets_value_when_variable_is_null, r#"string value = null; value ??= "set"; Console.WriteLine(value);"#, ["set"]);
csharp_case!(null_coalescing_assignment_keeps_existing_non_null_value, r#"string value = "keep"; value ??= "set"; Console.WriteLine(value);"#, ["keep"]);
csharp_case!(nested_null_conditional_walks_through_property_chain, r#"class Address { public string City { get; set; } } class User { public Address Address { get; set; } } var user = new User { Address = new Address { City = "Paris" } }; Console.WriteLine(user?.Address?.City ?? "none");"#, ["Paris"]);
csharp_case!(nested_null_conditional_stops_when_intermediate_is_null, r#"class Address { public string City { get; set; } } class User { public Address Address { get; set; } } var user = new User(); Console.WriteLine(user?.Address?.City ?? "none");"#, ["none"]);
csharp_case!(null_conditional_indexer_reads_existing_element, r#"int[] values = { 3, 4, 5 }; Console.WriteLine(values?[1] ?? -1);"#, ["4"]);
csharp_case!(null_conditional_indexer_uses_fallback_for_null_array, r#"int[] values = null; Console.WriteLine(values?[0] ?? -1);"#, ["-1"]);
csharp_case!(nullable_addition_uses_coalesced_default_when_missing, r#"int? left = null; int? right = 5; Console.WriteLine((left ?? 0) + (right ?? 0));"#, ["5"]);
csharp_case!(nullable_addition_uses_both_values_when_present, r#"int? left = 2; int? right = 5; Console.WriteLine((left ?? 0) + (right ?? 0));"#, ["7"]);
csharp_case!(delegate_null_conditional_invocation_skips_when_delegate_is_null, r#"using System; Action action = null; action?.Invoke(); Console.WriteLine("done");"#, ["done"]);
csharp_case!(delegate_null_conditional_invocation_calls_delegate_when_present, r#"using System; Action action = () => Console.WriteLine("ran"); action?.Invoke();"#, ["ran"]);
csharp_case!(null_coalescing_can_select_new_object_instance, r#"class Box { public string Name; } Box box = null; box ??= new Box { Name = "created" }; Console.WriteLine(box.Name);"#, ["created"]);
csharp_case!(null_conditional_length_returns_nullable_int_value, r#"string text = "four"; Console.WriteLine(text?.Length ?? 0);"#, ["4"]);
csharp_case!(null_conditional_length_returns_zero_for_null_string, r#"string text = null; Console.WriteLine(text?.Length ?? 0);"#, ["0"]);
csharp_case!(coalescing_chain_selects_first_non_null_candidate, r#"string first = null; string second = "B"; string third = "C"; Console.WriteLine(first ?? second ?? third);"#, ["B"]);