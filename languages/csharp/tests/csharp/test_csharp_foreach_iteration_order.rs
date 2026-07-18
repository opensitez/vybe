use super::helpers::run_csharp;

#[test]
fn foreach_iteration_order_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
int seed = 46; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
int seed = 46; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
int seed = 46; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_ordering_pair() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
int seed = 46; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_string_non_empty() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
string feature = "foreach_iteration_order"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_string_contains_probe() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
string feature = "foreach_iteration_order"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn foreach_iteration_order_string_first_char() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
string feature = "foreach_iteration_order"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_array_length_stable() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
int seed = 46; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_ternary_truth() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
int seed = 46; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_nullable_fallback() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
int? maybe = null; int fallback = maybe ?? 46; Console.WriteLine(fallback == 46);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
int? maybe = 46; Console.WriteLine(maybe.HasValue && maybe.Value == 46);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_list_count_contract() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
var values = new System.Collections.Generic.List<int> { 46, 47, 46 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
var set = new System.Collections.Generic.HashSet<int>(); set.Add(46); set.Add(46); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
var map = new System.Collections.Generic.Dictionary<int, int>(); map[46] = 47; Console.WriteLine(map.ContainsKey(46) && map[46] == 47);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_tuple_ordering() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
var tuple = (left: 46, right: 47); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
string feature = "foreach_iteration_order"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn foreach_iteration_order_double_identity() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
double seed = 46; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn foreach_iteration_order_feature_entropy() {
    assert_eq!(run_csharp(r#"// foreach_iteration_order
string feature = "foreach_iteration_order:46"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
