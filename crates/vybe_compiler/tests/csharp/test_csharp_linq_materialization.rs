//! LINQ operators that execute immediately (`ToList`, `Count`, aggregates).
use super::helpers::run_csharp;

#[test]
fn linq_to_list_materializes_query_before_source_mutation() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
using System.Linq;
var source = new List<int> { 1, 2 };
var snapshot = source.Select(x => x).ToList();
source.Add(3);
Console.WriteLine(snapshot.Count);
Console.WriteLine(source.Count);
"#
        ),
        &["2", "3"]
    );
}

#[test]
fn linq_count_executes_predicate_immediately() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
int checks = 0;
int total = new[] { 1, 2, 3 }.Count(x => { checks++; return x > 1; });
Console.WriteLine(total);
Console.WriteLine(checks);
"#
        ),
        &["2", "3"]
    );
}

#[test]
fn linq_sum_reduces_sequence_to_single_accumulated_value() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
Console.WriteLine(new[] { 1, 2, 3, 4 }.Sum());
"#
        ),
        &["10"]
    );
}

#[test]
fn linq_max_returns_greatest_element_by_default_comparer() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
Console.WriteLine(new[] { 3, 9, 4 }.Max());
"#
        ),
        &["9"]
    );
}

#[test]
fn linq_min_returns_smallest_element_by_default_comparer() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
Console.WriteLine(new[] { 3, 9, 4 }.Min());
"#
        ),
        &["3"]
    );
}

#[test]
fn linq_average_computes_mean_of_numeric_sequence() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
Console.WriteLine(new[] { 2, 4, 6 }.Average());
"#
        ),
        &["4"]
    );
}

#[test]
fn linq_first_throws_on_empty_sequence_when_not_caught() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
try {
    Console.WriteLine(new int[0].First());
} catch (System.InvalidOperationException) {
    Console.WriteLine("empty");
}
"#
        ),
        &["empty"]
    );
}

#[test]
fn linq_single_requires_exactly_one_element() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
Console.WriteLine(new[] { 42 }.Single());
"#
        ),
        &["42"]
    );
}

#[test]
fn linq_to_array_copies_elements_into_new_array() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
var copy = new[] { 5, 6 }.Select(x => x).ToArray();
Console.WriteLine(copy[1]);
"#
        ),
        &["6"]
    );
}

#[test]
fn linq_contains_searches_materialized_values_without_lazy_reenumeration() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
var data = new[] { "a", "b" };
Console.WriteLine(data.Contains("b"));
"#
        ),
        &["True"]
    );
}

#[test]
fn linq_reverse_materializes_reversed_order_in_new_sequence() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
foreach (var value in new[] { 1, 2, 3 }.Reverse()) Console.WriteLine(value);
"#
        ),
        &["3", "2", "1"]
    );
}

#[test]
fn linq_distinct_to_list_collapses_duplicates_during_materialization() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
var unique = new[] { 1, 1, 2, 2, 3 }.Distinct().ToList();
Console.WriteLine(unique.Count);
"#
        ),
        &["3"]
    );
}
