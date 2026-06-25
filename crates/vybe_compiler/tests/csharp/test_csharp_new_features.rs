//! Modern C# features: target-typed new, `nameof`, raw string literals, file-scoped namespace.
use super::helpers::run_csharp;

#[test]
fn target_typed_new_infers_list_type_from_variable() {
    assert_eq!(
        run_csharp(
            r#"System.Collections.Generic.List<int> nums = new();
nums.Add(1); nums.Add(2);
Console.WriteLine(nums.Count);"#
        ),
        &["2"]
    );
}

#[test]
fn target_typed_new_in_constructor_argument() {
    assert_eq!(
        run_csharp(
            r#"class Box { public System.Collections.Generic.List<int> Items; public Box(System.Collections.Generic.List<int> i){Items=i;} }
var b = new Box(new());
b.Items.Add(9);
Console.WriteLine(b.Items.Count);"#
        ),
        &["1"]
    );
}

#[test]
fn nameof_returns_string_name_of_variable() {
    assert_eq!(
        run_csharp(
            r#"int myCounter = 0;
Console.WriteLine(nameof(myCounter));"#
        ),
        &["myCounter"]
    );
}

#[test]
fn nameof_on_type_member_returns_member_name() {
    assert_eq!(
        run_csharp(
            r#"class Widget { public int Count; }
Console.WriteLine(nameof(Widget.Count));"#
        ),
        &["Count"]
    );
}

#[test]
fn conditional_ref_var_skips_copy_of_large_struct() {
    assert_eq!(
        run_csharp(
            r#"int[] arr = {1,2,3};
ref int val = ref arr[1];
val = 99;
Console.WriteLine(arr[1]);"#
        ),
        &["99"]
    );
}

#[test]
fn using_static_imports_type_members_without_qualifier() {
    assert_eq!(
        run_csharp(
            r#"using static System.Math;
Console.WriteLine(Sqrt(16));"#
        ),
        &["4"]
    );
}

#[test]
fn implicit_usings_allow_console_without_explicit_using() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(42);"#),
        &["42"]
    );
}

#[test]
fn range_from_end_produces_slice_of_array() {
    assert_eq!(
        run_csharp(
            r#"int[] arr = {1,2,3,4,5};
var last2 = arr[^2..];
Console.WriteLine(last2[0]); Console.WriteLine(last2[1]);"#
        ),
        &["4", "5"]
    );
}

#[test]
fn index_from_end_one_reads_last_element() {
    assert_eq!(
        run_csharp(
            r#"int[] arr = {10,20,30};
Console.WriteLine(arr[^1]);"#
        ),
        &["30"]
    );
}
