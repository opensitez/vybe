//! LINQ aggregate and quantifier operators.
use super::helpers::run_csharp;

#[test]
fn sum_adds_all_integers_in_sequence() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{1,2,3,4}.Sum());"#),
        &["10"]
    );
}

#[test]
fn count_without_predicate_returns_element_count() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{1,2,3}.Count());"#),
        &["3"]
    );
}

#[test]
fn count_with_predicate_counts_matching_elements() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{1,2,3,4,5}.Count(x => x%2==0));"#),
        &["2"]
    );
}

#[test]
fn min_returns_smallest_element() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{5,1,9,3}.Min());"#),
        &["1"]
    );
}

#[test]
fn max_returns_largest_element() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{5,1,9,3}.Max());"#),
        &["9"]
    );
}

#[test]
fn average_returns_mean_as_double() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{1,2,3}.Average());"#),
        &["2"]
    );
}

#[test]
fn any_returns_true_when_predicate_satisfied() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{1,2,3}.Any(x => x>2));"#),
        &["True"]
    );
}

#[test]
fn any_returns_false_on_empty_sequence() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(System.Array.Empty<int>().Any());"#),
        &["False"]
    );
}

#[test]
fn all_returns_false_when_one_element_fails_predicate() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{2,4,5}.All(x => x%2==0));"#),
        &["False"]
    );
}

#[test]
fn first_returns_first_element_of_sequence() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{10,20,30}.First());"#),
        &["10"]
    );
}

#[test]
fn first_with_predicate_skips_to_matching_element() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{1,2,3,4}.First(x => x>2));"#),
        &["3"]
    );
}

#[test]
fn last_returns_final_element_of_sequence() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{10,20,30}.Last());"#),
        &["30"]
    );
}

#[test]
fn single_throws_when_sequence_has_more_than_one_match() {
    assert_eq!(
        run_csharp(
            r#"
string result = "ok";
try { new[]{1,2}.Single(); }
catch(System.InvalidOperationException) { result = "many"; }
Console.WriteLine(result);"#
        ),
        &["many"]
    );
}

#[test]
fn aggregate_folds_sequence_with_seed_and_accumulator() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{1,2,3,4}.Aggregate(0, (acc, x) => acc + x));"#),
        &["10"]
    );
}

#[test]
fn contains_returns_true_for_present_value_in_sequence() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{1,2,3}.Contains(2));"#),
        &["True"]
    );
}
