//! LINQ to Objects deferred execution: queries do not run until enumeration,
//! and re-enumeration re-executes the pipeline against the current source.
use super::helpers::run_csharp;

#[test]
fn linq_select_does_not_run_until_foreach_starts() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
int sideEffects = 0;
var query = new[] { 1, 2 }.Select(x => { sideEffects++; return x; });
Console.WriteLine(sideEffects);
foreach (var _ in query) { }
Console.WriteLine(sideEffects);
"#
        ),
        &["0", "2"]
    );
}

#[test]
fn linq_where_filter_runs_only_during_enumeration() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
int checks = 0;
var query = new[] { 1, 2, 3, 4 }.Where(x => { checks++; return x % 2 == 0; });
Console.WriteLine(checks);
foreach (var value in query) Console.WriteLine(value);
Console.WriteLine(checks);
"#
        ),
        &["0", "2", "4", "4"]
    );
}

#[test]
fn linq_pipeline_mutating_source_before_enumeration_sees_new_items() {
    assert_eq!(
        run_csharp(
            r#"
using System.Collections.Generic;
using System.Linq;
var data = new List<int> { 1, 2 };
var query = data.Select(x => x * 10);
data.Add(3);
foreach (var value in query) Console.WriteLine(value);
"#
        ),
        &["10", "20", "30"]
    );
}

#[test]
fn linq_second_enumeration_reexecutes_select_projection() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
int projections = 0;
var query = new[] { 5 }.Select(x => { projections++; return x + 1; });
Console.WriteLine(query.First());
Console.WriteLine(projections);
Console.WriteLine(query.First());
Console.WriteLine(projections);
"#
        ),
        &["6", "1", "6", "2"]
    );
}

#[test]
fn linq_orderby_defers_sort_until_materialization() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
int comparisons = 0;
var query = new[] { 3, 1, 2 }.OrderBy(x => { comparisons++; return x; });
Console.WriteLine(comparisons);
foreach (var value in query) Console.WriteLine(value);
"#
        ),
        &["0", "1", "2", "3"]
    );
}

#[test]
fn linq_take_short_circuits_without_visiting_entire_source() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
int visited = 0;
var query = Enumerable.Range(1, 100).Select(x => { visited++; return x; }).Take(2);
foreach (var value in query) Console.WriteLine(value);
Console.WriteLine(visited);
"#
        ),
        &["1", "2", "2"]
    );
}

#[test]
fn linq_skip_defers_discard_until_enumeration() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
var query = new[] { 10, 20, 30, 40 }.Skip(2);
foreach (var value in query) Console.WriteLine(value);
"#
        ),
        &["30", "40"]
    );
}

#[test]
fn linq_select_many_flattens_nested_sequences_lazily() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
var query = new[] { "ab", "c" }.SelectMany(word => word.Select(ch => ch));
foreach (var ch in query) Console.WriteLine(ch);
"#
        ),
        &["a", "b", "c"]
    );
}

#[test]
fn linq_aggregate_eagerly_reduces_without_intermediate_list() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
int sum = new[] { 1, 2, 3, 4 }.Aggregate(0, (acc, x) => acc + x);
Console.WriteLine(sum);
"#
        ),
        &["10"]
    );
}

#[test]
fn linq_any_short_circuits_on_first_matching_element() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
int probes = 0;
bool found = new[] { 1, 2, 3 }.Any(x => { probes++; return x == 2; });
Console.WriteLine(found);
Console.WriteLine(probes);
"#
        ),
        &["True", "2"]
    );
}

#[test]
fn linq_all_short_circuits_on_first_failing_element() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
int probes = 0;
bool ok = new[] { 2, 4, 5, 8 }.All(x => { probes++; return x % 2 == 0; });
Console.WriteLine(ok);
Console.WriteLine(probes);
"#
        ),
        &["False", "3"]
    );
}

#[test]
fn linq_distinct_uses_default_equality_comparer_lazily() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
foreach (var value in new[] { 1, 1, 2, 2, 3 }.Distinct()) Console.WriteLine(value);
"#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn linq_zip_pairs_elements_until_shorter_sequence_ends() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
foreach (var pair in new[] { 1, 2, 3 }.Zip(new[] { 10, 20 }, (a, b) => a + b)) Console.WriteLine(pair);
"#
        ),
        &["11", "22"]
    );
}

#[test]
fn linq_oftype_filters_runtime_types_during_enumeration() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
object[] items = { 1, "a", 2, "b", 3 };
foreach (var text in items.OfType<string>()) Console.WriteLine(text);
"#
        ),
        &["a", "b"]
    );
}

#[test]
fn linq_cast_unboxes_numeric_sequence_to_int_stream() {
    assert_eq!(
        run_csharp(
            r#"
using System.Linq;
object[] boxed = { 1, 2, 3 };
foreach (var value in boxed.Cast<int>()) Console.WriteLine(value + 1);
"#
        ),
        &["2", "3", "4"]
    );
}
