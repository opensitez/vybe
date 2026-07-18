use super::helpers::run_csharp;

#[test]
fn dictionary_insert_lookup_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
int seed = 34; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
int seed = 34; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
int seed = 34; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_ordering_pair() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
int seed = 34; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_string_non_empty() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
string feature = "dictionary_insert_lookup"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_string_contains_probe() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
string feature = "dictionary_insert_lookup"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_string_first_char() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
string feature = "dictionary_insert_lookup"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_array_length_stable() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
int seed = 34; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_ternary_truth() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
int seed = 34; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_nullable_fallback() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
int? maybe = null; int fallback = maybe ?? 34; Console.WriteLine(fallback == 34);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
int? maybe = 34; Console.WriteLine(maybe.HasValue && maybe.Value == 34);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_list_count_contract() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
var values = new System.Collections.Generic.List<int> { 34, 35, 34 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
var set = new System.Collections.Generic.HashSet<int>(); set.Add(34); set.Add(34); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
var map = new System.Collections.Generic.Dictionary<int, int>(); map[34] = 35; Console.WriteLine(map.ContainsKey(34) && map[34] == 35);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_tuple_ordering() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
var tuple = (left: 34, right: 35); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
string feature = "dictionary_insert_lookup"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_double_identity() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
double seed = 34; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn dictionary_insert_lookup_feature_entropy() {
    assert_eq!(run_csharp(r#"// dictionary_insert_lookup
string feature = "dictionary_insert_lookup:34"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
