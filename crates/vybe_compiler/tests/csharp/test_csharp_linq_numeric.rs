//! LINQ on numeric sequences: Sum, Average, Min, Max with complex selectors.
use super::helpers::run_csharp;

#[test]
fn sum_with_selector_projects_before_summing() {
    assert_eq!(
        run_csharp(
            r#"var words=new[]{"hello","world","foo"};
Console.WriteLine(words.Sum(w=>w.Length));"#
        ),
        &["13"]
    );
}

#[test]
fn average_of_integer_sequence_returns_double() {
    assert_eq!(
        run_csharp(
            r#"double avg=new[]{1,2,3,4,5}.Average();
Console.WriteLine(avg);"#
        ),
        &["3"]
    );
}

#[test]
fn min_with_custom_selector() {
    assert_eq!(
        run_csharp(
            r#"var words=new[]{"cat","elephant","ox"};
Console.WriteLine(words.Min(w=>w.Length));"#
        ),
        &["2"]
    );
}

#[test]
fn max_by_returns_whole_element_not_just_key() {
    assert_eq!(
        run_csharp(
            r#"var words=new[]{"cat","elephant","ox"};
Console.WriteLine(words.MaxBy(w=>w.Length));"#
        ),
        &["elephant"]
    );
}

#[test]
fn sum_of_empty_sequence_is_zero() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Array.Empty<int>().Sum());"#),
        &["0"]
    );
}

#[test]
fn count_with_predicate_counts_matching() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{1,2,3,4,5,6}.Count(n=>n%2==0));"#),
        &["3"]
    );
}

#[test]
fn long_count_works_on_large_range() {
    assert_eq!(
        run_csharp(
            r#"long c=Enumerable.Range(0,1000).LongCount();
Console.WriteLine(c);"#
        ),
        &["1000"]
    );
}
