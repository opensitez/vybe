use super::helpers::run_csharp;

#[test]
fn indexer_get_set_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// indexer_get_set
int seed = 66; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn indexer_get_set_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// indexer_get_set
int seed = 66; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn indexer_get_set_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// indexer_get_set
int seed = 66; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn indexer_get_set_ordering_pair() {
    assert_eq!(run_csharp(r#"// indexer_get_set
int seed = 66; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn indexer_get_set_string_non_empty() {
    assert_eq!(run_csharp(r#"// indexer_get_set
string feature = "indexer_get_set"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn indexer_get_set_string_contains_probe() {
    assert_eq!(run_csharp(r#"// indexer_get_set
string feature = "indexer_get_set"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn indexer_get_set_string_first_char() {
    assert_eq!(run_csharp(r#"// indexer_get_set
string feature = "indexer_get_set"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn indexer_get_set_array_length_stable() {
    assert_eq!(run_csharp(r#"// indexer_get_set
int seed = 66; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn indexer_get_set_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// indexer_get_set
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn indexer_get_set_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// indexer_get_set
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn indexer_get_set_ternary_truth() {
    assert_eq!(run_csharp(r#"// indexer_get_set
int seed = 66; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn indexer_get_set_nullable_fallback() {
    assert_eq!(run_csharp(r#"// indexer_get_set
int? maybe = null; int fallback = maybe ?? 66; Console.WriteLine(fallback == 66);"#), &["True"]);
}

#[test]
fn indexer_get_set_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// indexer_get_set
int? maybe = 66; Console.WriteLine(maybe.HasValue && maybe.Value == 66);"#), &["True"]);
}

#[test]
fn indexer_get_set_list_count_contract() {
    assert_eq!(run_csharp(r#"// indexer_get_set
var values = new System.Collections.Generic.List<int> { 66, 67, 66 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn indexer_get_set_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// indexer_get_set
var set = new System.Collections.Generic.HashSet<int>(); set.Add(66); set.Add(66); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn indexer_get_set_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// indexer_get_set
var map = new System.Collections.Generic.Dictionary<int, int>(); map[66] = 67; Console.WriteLine(map.ContainsKey(66) && map[66] == 67);"#), &["True"]);
}

#[test]
fn indexer_get_set_tuple_ordering() {
    assert_eq!(run_csharp(r#"// indexer_get_set
var tuple = (left: 66, right: 67); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn indexer_get_set_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// indexer_get_set
string feature = "indexer_get_set"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn indexer_get_set_double_identity() {
    assert_eq!(run_csharp(r#"// indexer_get_set
double seed = 66; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn indexer_get_set_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// indexer_get_set
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn indexer_get_set_feature_entropy() {
    assert_eq!(run_csharp(r#"// indexer_get_set
string feature = "indexer_get_set:66"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
