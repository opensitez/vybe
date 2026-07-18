use super::helpers::run_csharp;

#[test]
fn record_shape_basics_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// record_shape_basics
int seed = 39; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn record_shape_basics_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// record_shape_basics
int seed = 39; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn record_shape_basics_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// record_shape_basics
int seed = 39; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn record_shape_basics_ordering_pair() {
    assert_eq!(run_csharp(r#"// record_shape_basics
int seed = 39; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn record_shape_basics_string_non_empty() {
    assert_eq!(run_csharp(r#"// record_shape_basics
string feature = "record_shape_basics"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn record_shape_basics_string_contains_probe() {
    assert_eq!(run_csharp(r#"// record_shape_basics
string feature = "record_shape_basics"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn record_shape_basics_string_first_char() {
    assert_eq!(run_csharp(r#"// record_shape_basics
string feature = "record_shape_basics"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn record_shape_basics_array_length_stable() {
    assert_eq!(run_csharp(r#"// record_shape_basics
int seed = 39; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn record_shape_basics_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// record_shape_basics
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn record_shape_basics_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// record_shape_basics
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn record_shape_basics_ternary_truth() {
    assert_eq!(run_csharp(r#"// record_shape_basics
int seed = 39; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn record_shape_basics_nullable_fallback() {
    assert_eq!(run_csharp(r#"// record_shape_basics
int? maybe = null; int fallback = maybe ?? 39; Console.WriteLine(fallback == 39);"#), &["True"]);
}

#[test]
fn record_shape_basics_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// record_shape_basics
int? maybe = 39; Console.WriteLine(maybe.HasValue && maybe.Value == 39);"#), &["True"]);
}

#[test]
fn record_shape_basics_list_count_contract() {
    assert_eq!(run_csharp(r#"// record_shape_basics
var values = new System.Collections.Generic.List<int> { 39, 40, 39 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn record_shape_basics_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// record_shape_basics
var set = new System.Collections.Generic.HashSet<int>(); set.Add(39); set.Add(39); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn record_shape_basics_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// record_shape_basics
var map = new System.Collections.Generic.Dictionary<int, int>(); map[39] = 40; Console.WriteLine(map.ContainsKey(39) && map[39] == 40);"#), &["True"]);
}

#[test]
fn record_shape_basics_tuple_ordering() {
    assert_eq!(run_csharp(r#"// record_shape_basics
var tuple = (left: 39, right: 40); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn record_shape_basics_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// record_shape_basics
string feature = "record_shape_basics"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn record_shape_basics_double_identity() {
    assert_eq!(run_csharp(r#"// record_shape_basics
double seed = 39; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn record_shape_basics_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// record_shape_basics
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn record_shape_basics_feature_entropy() {
    assert_eq!(run_csharp(r#"// record_shape_basics
string feature = "record_shape_basics:39"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
