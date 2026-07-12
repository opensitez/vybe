//! `ref` returns alias caller storage; mutations flow back into the source.
use super::helpers::run_csharp;

#[test]
fn ref_return_allows_mutating_array_element_through_alias() {
    assert_eq!(
        run_csharp(
            r#"
int[] data = { 1, 2, 3 };
ref int Slot(int index) => ref data[index];
ref int cell = ref Slot(1);
cell = 9;
Console.WriteLine(data[1]);
"#
        ),
        &["9"]
    );
}

#[test]
fn ref_return_from_local_function_updates_outer_variable() {
    assert_eq!(
        run_csharp(
            r#"
int total = 5;
ref int Bump() => ref total;
ref int view = ref Bump();
view += 2;
Console.WriteLine(total);
"#
        ),
        &["7"]
    );
}

#[test]
fn ref_return_chains_to_second_ref_local_without_copying_value() {
    assert_eq!(
        run_csharp(
            r#"
int[] values = { 10, 20 };
ref int First() => ref values[0];
ref int alias = ref First();
alias = 99;
Console.WriteLine(values[0]);
"#
        ),
        &["99"]
    );
}
