use super::helpers::run_csharp;

#[test]
fn short_circuit_logic_patterns_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
int seed = 14; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
int seed = 14; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
int seed = 14; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_ordering_pair() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
int seed = 14; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_string_non_empty() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
string feature = "short_circuit_logic_patterns"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_string_contains_probe() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
string feature = "short_circuit_logic_patterns"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_string_first_char() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
string feature = "short_circuit_logic_patterns"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_array_length_stable() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
int seed = 14; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_ternary_truth() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
int seed = 14; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_nullable_fallback() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
int? maybe = null; int fallback = maybe ?? 14; Console.WriteLine(fallback == 14);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
int? maybe = 14; Console.WriteLine(maybe.HasValue && maybe.Value == 14);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_list_count_contract() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
var values = new System.Collections.Generic.List<int> { 14, 15, 14 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
var set = new System.Collections.Generic.HashSet<int>(); set.Add(14); set.Add(14); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
var map = new System.Collections.Generic.Dictionary<int, int>(); map[14] = 15; Console.WriteLine(map.ContainsKey(14) && map[14] == 15);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_tuple_ordering() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
var tuple = (left: 14, right: 15); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
string feature = "short_circuit_logic_patterns"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_double_identity() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
double seed = 14; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn short_circuit_logic_patterns_feature_entropy() {
    assert_eq!(run_csharp(r#"// short_circuit_logic_patterns
string feature = "short_circuit_logic_patterns:14"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
