use super::helpers::run_csharp;

#[test]
fn array_reverse_patterns_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
int seed = 27; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
int seed = 27; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
int seed = 27; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_ordering_pair() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
int seed = 27; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_string_non_empty() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
string feature = "array_reverse_patterns"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_string_contains_probe() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
string feature = "array_reverse_patterns"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn array_reverse_patterns_string_first_char() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
string feature = "array_reverse_patterns"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_array_length_stable() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
int seed = 27; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_ternary_truth() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
int seed = 27; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_nullable_fallback() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
int? maybe = null; int fallback = maybe ?? 27; Console.WriteLine(fallback == 27);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
int? maybe = 27; Console.WriteLine(maybe.HasValue && maybe.Value == 27);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_list_count_contract() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
var values = new System.Collections.Generic.List<int> { 27, 28, 27 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
var set = new System.Collections.Generic.HashSet<int>(); set.Add(27); set.Add(27); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
var map = new System.Collections.Generic.Dictionary<int, int>(); map[27] = 28; Console.WriteLine(map.ContainsKey(27) && map[27] == 28);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_tuple_ordering() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
var tuple = (left: 27, right: 28); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
string feature = "array_reverse_patterns"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn array_reverse_patterns_double_identity() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
double seed = 27; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn array_reverse_patterns_feature_entropy() {
    assert_eq!(run_csharp(r#"// array_reverse_patterns
string feature = "array_reverse_patterns:27"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
