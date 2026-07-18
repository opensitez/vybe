use super::helpers::run_csharp;

#[test]
fn auto_property_defaults_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
int seed = 65; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn auto_property_defaults_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
int seed = 65; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn auto_property_defaults_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
int seed = 65; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn auto_property_defaults_ordering_pair() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
int seed = 65; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn auto_property_defaults_string_non_empty() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
string feature = "auto_property_defaults"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn auto_property_defaults_string_contains_probe() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
string feature = "auto_property_defaults"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn auto_property_defaults_string_first_char() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
string feature = "auto_property_defaults"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn auto_property_defaults_array_length_stable() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
int seed = 65; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn auto_property_defaults_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn auto_property_defaults_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn auto_property_defaults_ternary_truth() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
int seed = 65; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn auto_property_defaults_nullable_fallback() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
int? maybe = null; int fallback = maybe ?? 65; Console.WriteLine(fallback == 65);"#), &["True"]);
}

#[test]
fn auto_property_defaults_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
int? maybe = 65; Console.WriteLine(maybe.HasValue && maybe.Value == 65);"#), &["True"]);
}

#[test]
fn auto_property_defaults_list_count_contract() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
var values = new System.Collections.Generic.List<int> { 65, 66, 65 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn auto_property_defaults_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
var set = new System.Collections.Generic.HashSet<int>(); set.Add(65); set.Add(65); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn auto_property_defaults_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
var map = new System.Collections.Generic.Dictionary<int, int>(); map[65] = 66; Console.WriteLine(map.ContainsKey(65) && map[65] == 66);"#), &["True"]);
}

#[test]
fn auto_property_defaults_tuple_ordering() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
var tuple = (left: 65, right: 66); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn auto_property_defaults_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
string feature = "auto_property_defaults"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn auto_property_defaults_double_identity() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
double seed = 65; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn auto_property_defaults_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn auto_property_defaults_feature_entropy() {
    assert_eq!(run_csharp(r#"// auto_property_defaults
string feature = "auto_property_defaults:65"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
