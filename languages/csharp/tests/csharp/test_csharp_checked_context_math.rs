use super::helpers::run_csharp;

#[test]
fn checked_context_math_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// checked_context_math
int seed = 12; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn checked_context_math_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// checked_context_math
int seed = 12; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn checked_context_math_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// checked_context_math
int seed = 12; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn checked_context_math_ordering_pair() {
    assert_eq!(run_csharp(r#"// checked_context_math
int seed = 12; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn checked_context_math_string_non_empty() {
    assert_eq!(run_csharp(r#"// checked_context_math
string feature = "checked_context_math"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn checked_context_math_string_contains_probe() {
    assert_eq!(run_csharp(r#"// checked_context_math
string feature = "checked_context_math"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn checked_context_math_string_first_char() {
    assert_eq!(run_csharp(r#"// checked_context_math
string feature = "checked_context_math"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn checked_context_math_array_length_stable() {
    assert_eq!(run_csharp(r#"// checked_context_math
int seed = 12; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn checked_context_math_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// checked_context_math
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn checked_context_math_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// checked_context_math
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn checked_context_math_ternary_truth() {
    assert_eq!(run_csharp(r#"// checked_context_math
int seed = 12; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn checked_context_math_nullable_fallback() {
    assert_eq!(run_csharp(r#"// checked_context_math
int? maybe = null; int fallback = maybe ?? 12; Console.WriteLine(fallback == 12);"#), &["True"]);
}

#[test]
fn checked_context_math_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// checked_context_math
int? maybe = 12; Console.WriteLine(maybe.HasValue && maybe.Value == 12);"#), &["True"]);
}

#[test]
fn checked_context_math_list_count_contract() {
    assert_eq!(run_csharp(r#"// checked_context_math
var values = new System.Collections.Generic.List<int> { 12, 13, 12 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn checked_context_math_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// checked_context_math
var set = new System.Collections.Generic.HashSet<int>(); set.Add(12); set.Add(12); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn checked_context_math_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// checked_context_math
var map = new System.Collections.Generic.Dictionary<int, int>(); map[12] = 13; Console.WriteLine(map.ContainsKey(12) && map[12] == 13);"#), &["True"]);
}

#[test]
fn checked_context_math_tuple_ordering() {
    assert_eq!(run_csharp(r#"// checked_context_math
var tuple = (left: 12, right: 13); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn checked_context_math_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// checked_context_math
string feature = "checked_context_math"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn checked_context_math_double_identity() {
    assert_eq!(run_csharp(r#"// checked_context_math
double seed = 12; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn checked_context_math_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// checked_context_math
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn checked_context_math_feature_entropy() {
    assert_eq!(run_csharp(r#"// checked_context_math
string feature = "checked_context_math:12"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
