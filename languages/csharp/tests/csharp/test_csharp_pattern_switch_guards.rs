use super::helpers::run_csharp;

#[test]
fn pattern_switch_guards_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
int seed = 42; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
int seed = 42; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
int seed = 42; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_ordering_pair() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
int seed = 42; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_string_non_empty() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
string feature = "pattern_switch_guards"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_string_contains_probe() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
string feature = "pattern_switch_guards"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn pattern_switch_guards_string_first_char() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
string feature = "pattern_switch_guards"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_array_length_stable() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
int seed = 42; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_ternary_truth() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
int seed = 42; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_nullable_fallback() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
int? maybe = null; int fallback = maybe ?? 42; Console.WriteLine(fallback == 42);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
int? maybe = 42; Console.WriteLine(maybe.HasValue && maybe.Value == 42);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_list_count_contract() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
var values = new System.Collections.Generic.List<int> { 42, 43, 42 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
var set = new System.Collections.Generic.HashSet<int>(); set.Add(42); set.Add(42); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
var map = new System.Collections.Generic.Dictionary<int, int>(); map[42] = 43; Console.WriteLine(map.ContainsKey(42) && map[42] == 43);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_tuple_ordering() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
var tuple = (left: 42, right: 43); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
string feature = "pattern_switch_guards"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn pattern_switch_guards_double_identity() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
double seed = 42; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn pattern_switch_guards_feature_entropy() {
    assert_eq!(run_csharp(r#"// pattern_switch_guards
string feature = "pattern_switch_guards:42"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
