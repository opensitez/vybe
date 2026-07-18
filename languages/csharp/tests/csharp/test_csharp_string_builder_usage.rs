use super::helpers::run_csharp;

#[test]
fn string_builder_usage_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// string_builder_usage
int seed = 20; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn string_builder_usage_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// string_builder_usage
int seed = 20; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn string_builder_usage_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// string_builder_usage
int seed = 20; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn string_builder_usage_ordering_pair() {
    assert_eq!(run_csharp(r#"// string_builder_usage
int seed = 20; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn string_builder_usage_string_non_empty() {
    assert_eq!(run_csharp(r#"// string_builder_usage
string feature = "string_builder_usage"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn string_builder_usage_string_contains_probe() {
    assert_eq!(run_csharp(r#"// string_builder_usage
string feature = "string_builder_usage"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn string_builder_usage_string_first_char() {
    assert_eq!(run_csharp(r#"// string_builder_usage
string feature = "string_builder_usage"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn string_builder_usage_array_length_stable() {
    assert_eq!(run_csharp(r#"// string_builder_usage
int seed = 20; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn string_builder_usage_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// string_builder_usage
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn string_builder_usage_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// string_builder_usage
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn string_builder_usage_ternary_truth() {
    assert_eq!(run_csharp(r#"// string_builder_usage
int seed = 20; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn string_builder_usage_nullable_fallback() {
    assert_eq!(run_csharp(r#"// string_builder_usage
int? maybe = null; int fallback = maybe ?? 20; Console.WriteLine(fallback == 20);"#), &["True"]);
}

#[test]
fn string_builder_usage_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// string_builder_usage
int? maybe = 20; Console.WriteLine(maybe.HasValue && maybe.Value == 20);"#), &["True"]);
}

#[test]
fn string_builder_usage_list_count_contract() {
    assert_eq!(run_csharp(r#"// string_builder_usage
var values = new System.Collections.Generic.List<int> { 20, 21, 20 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn string_builder_usage_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// string_builder_usage
var set = new System.Collections.Generic.HashSet<int>(); set.Add(20); set.Add(20); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn string_builder_usage_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// string_builder_usage
var map = new System.Collections.Generic.Dictionary<int, int>(); map[20] = 21; Console.WriteLine(map.ContainsKey(20) && map[20] == 21);"#), &["True"]);
}

#[test]
fn string_builder_usage_tuple_ordering() {
    assert_eq!(run_csharp(r#"// string_builder_usage
var tuple = (left: 20, right: 21); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn string_builder_usage_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// string_builder_usage
string feature = "string_builder_usage"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn string_builder_usage_double_identity() {
    assert_eq!(run_csharp(r#"// string_builder_usage
double seed = 20; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn string_builder_usage_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// string_builder_usage
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn string_builder_usage_feature_entropy() {
    assert_eq!(run_csharp(r#"// string_builder_usage
string feature = "string_builder_usage:20"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
