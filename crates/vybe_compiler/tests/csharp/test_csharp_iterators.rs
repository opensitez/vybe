//! Iterator methods with `yield return` and `yield break` semantics.
use super::helpers::run_csharp;

#[test]
fn yield_return_sequence_consumed_by_foreach() {
    assert_eq!(
        run_csharp(
            r#"System.Collections.Generic.IEnumerable<int> Gen() {
    yield return 1; yield return 2; yield return 3;
}
int sum = 0;
foreach(var n in Gen()) sum += n;
Console.WriteLine(sum);"#
        ),
        &["6"]
    );
}

#[test]
fn yield_break_stops_iteration_early() {
    assert_eq!(
        run_csharp(
            r#"System.Collections.Generic.IEnumerable<int> Gen() {
    yield return 1;
    yield break;
    yield return 2;
}
int count = 0;
foreach(var _ in Gen()) count++;
Console.WriteLine(count);"#
        ),
        &["1"]
    );
}

#[test]
fn yield_in_loop_produces_computed_sequence() {
    assert_eq!(
        run_csharp(
            r#"System.Collections.Generic.IEnumerable<int> Range(int n) {
    for(int i=0; i<n; i++) yield return i;
}
int sum=0;
foreach(var x in Range(5)) sum+=x;
Console.WriteLine(sum);"#
        ),
        &["10"]
    );
}

#[test]
fn yield_is_lazy_factory_not_calls_body_before_iteration() {
    assert_eq!(
        run_csharp(
            r#"int calls=0;
System.Collections.Generic.IEnumerable<int> Lazy() {
    calls++;
    yield return 1;
}
Console.WriteLine(calls);
var seq = Lazy();
Console.WriteLine(calls);
foreach(var _ in seq) {}
Console.WriteLine(calls);"#
        ),
        &["0", "0", "1"]
    );
}

#[test]
fn multiple_foreach_iterations_restart_the_iterator() {
    assert_eq!(
        run_csharp(
            r#"System.Collections.Generic.IEnumerable<int> Three() {
    yield return 1; yield return 2; yield return 3;
}
int total=0;
foreach(var x in Three()) total+=x;
foreach(var x in Three()) total+=x;
Console.WriteLine(total);"#
        ),
        &["12"]
    );
}

#[test]
fn iterator_over_string_chars_via_yield() {
    assert_eq!(
        run_csharp(
            r#"System.Collections.Generic.IEnumerable<char> Vowels(string s) {
    foreach(char c in s) if("aeiou".Contains(c)) yield return c;
}
int count=0;
foreach(var _ in Vowels("hello world")) count++;
Console.WriteLine(count);"#
        ),
        &["3"]
    );
}
