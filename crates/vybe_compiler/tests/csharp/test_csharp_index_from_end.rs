//! Index-from-end (`^n`) addresses elements relative to sequence length.
use super::helpers::run_csharp;

#[test]
fn index_from_end_one_reads_last_element_of_array() {
    assert_eq!(
        run_csharp(
            r#"
int[] data = { 10, 20, 30 };
Console.WriteLine(data[^1]);
"#
        ),
        &["30"]
    );
}

#[test]
fn range_to_end_from_index_from_end_produces_tail_slice() {
    assert_eq!(
        run_csharp(
            r#"
int[] data = { 1, 2, 3, 4 };
var tail = data[2..^0];
Console.WriteLine(tail.Length);
Console.WriteLine(tail[0]);
"#
        ),
        &["2", "3"]
    );
}
