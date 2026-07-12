//! LINQ set operations: Distinct, Union, Intersect, Except, SequenceEqual.
use super::helpers::run_csharp;

#[test]
fn distinct_removes_duplicate_elements() {
    assert_eq!(
        run_csharp(
            r#"var result = new[]{1,2,2,3,1}.Distinct().OrderBy(x=>x);
foreach(var x in result) Console.WriteLine(x);"#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn union_merges_two_sequences_without_duplicates() {
    assert_eq!(
        run_csharp(
            r#"var result = new[]{1,2,3}.Union(new[]{3,4,5}).OrderBy(x=>x);
foreach(var x in result) Console.WriteLine(x);"#
        ),
        &["1", "2", "3", "4", "5"]
    );
}

#[test]
fn intersect_yields_elements_present_in_both_sequences() {
    assert_eq!(
        run_csharp(
            r#"var result = new[]{1,2,3,4}.Intersect(new[]{2,4,6}).OrderBy(x=>x);
foreach(var x in result) Console.WriteLine(x);"#
        ),
        &["2", "4"]
    );
}

#[test]
fn except_yields_elements_in_first_not_in_second() {
    assert_eq!(
        run_csharp(
            r#"var result = new[]{1,2,3,4}.Except(new[]{2,4}).OrderBy(x=>x);
foreach(var x in result) Console.WriteLine(x);"#
        ),
        &["1", "3"]
    );
}

#[test]
fn sequence_equal_returns_true_for_matching_sequences() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{1,2,3}.SequenceEqual(new[]{1,2,3}));"#),
        &["True"]
    );
}

#[test]
fn sequence_equal_returns_false_for_different_order() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{1,2,3}.SequenceEqual(new[]{3,2,1}));"#),
        &["False"]
    );
}

#[test]
fn concat_chains_two_sequences_preserving_all_elements() {
    assert_eq!(
        run_csharp(
            r#"var result = new[]{1,2}.Concat(new[]{3,4});
Console.WriteLine(result.Count());"#
        ),
        &["4"]
    );
}

#[test]
fn zip_pairs_elements_from_two_sequences_by_position() {
    assert_eq!(
        run_csharp(
            r#"var result = new[]{1,2,3}.Zip(new[]{10,20,30}, (a,b) => a*b);
foreach(var x in result) Console.WriteLine(x);"#
        ),
        &["10", "40", "90"]
    );
}
