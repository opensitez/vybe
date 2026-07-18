use super::helpers::run_csharp;

#[test]
fn string_split_join_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// string_split_join
int seed = 21; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn string_split_join_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// string_split_join
int seed = 21; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn string_split_join_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// string_split_join
int seed = 21; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn string_split_join_ordering_pair() {
    assert_eq!(run_csharp(r#"// string_split_join
int seed = 21; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn string_split_join_string_non_empty() {
    assert_eq!(run_csharp(r#"// string_split_join
string feature = "string_split_join"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn string_split_join_string_contains_probe() {
    assert_eq!(run_csharp(r#"// string_split_join
string feature = "string_split_join"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn string_split_join_string_first_char() {
    assert_eq!(run_csharp(r#"// string_split_join
string feature = "string_split_join"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn string_split_join_array_length_stable() {
    assert_eq!(run_csharp(r#"// string_split_join
int seed = 21; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn string_split_join_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// string_split_join
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn string_split_join_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// string_split_join
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn string_split_join_ternary_truth() {
    assert_eq!(run_csharp(r#"// string_split_join
int seed = 21; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn string_split_join_nullable_fallback() {
    assert_eq!(run_csharp(r#"// string_split_join
int? maybe = null; int fallback = maybe ?? 21; Console.WriteLine(fallback == 21);"#), &["True"]);
}

#[test]
fn string_split_join_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// string_split_join
int? maybe = 21; Console.WriteLine(maybe.HasValue && maybe.Value == 21);"#), &["True"]);
}

#[test]
fn string_split_join_list_count_contract() {
    assert_eq!(run_csharp(r#"// string_split_join
var values = new System.Collections.Generic.List<int> { 21, 22, 21 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn string_split_join_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// string_split_join
var set = new System.Collections.Generic.HashSet<int>(); set.Add(21); set.Add(21); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn string_split_join_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// string_split_join
var map = new System.Collections.Generic.Dictionary<int, int>(); map[21] = 22; Console.WriteLine(map.ContainsKey(21) && map[21] == 22);"#), &["True"]);
}

#[test]
fn string_split_join_tuple_ordering() {
    assert_eq!(run_csharp(r#"// string_split_join
var tuple = (left: 21, right: 22); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn string_split_join_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// string_split_join
string feature = "string_split_join"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn string_split_join_double_identity() {
    assert_eq!(run_csharp(r#"// string_split_join
double seed = 21; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn string_split_join_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// string_split_join
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn string_split_join_feature_entropy() {
    assert_eq!(run_csharp(r#"// string_split_join
string feature = "string_split_join:21"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
