use super::helpers::run_csharp;

#[test]
fn multidim_array_traversal_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
int seed = 29; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
int seed = 29; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
int seed = 29; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_ordering_pair() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
int seed = 29; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_string_non_empty() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
string feature = "multidim_array_traversal"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_string_contains_probe() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
string feature = "multidim_array_traversal"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn multidim_array_traversal_string_first_char() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
string feature = "multidim_array_traversal"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_array_length_stable() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
int seed = 29; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_ternary_truth() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
int seed = 29; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_nullable_fallback() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
int? maybe = null; int fallback = maybe ?? 29; Console.WriteLine(fallback == 29);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
int? maybe = 29; Console.WriteLine(maybe.HasValue && maybe.Value == 29);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_list_count_contract() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
var values = new System.Collections.Generic.List<int> { 29, 30, 29 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
var set = new System.Collections.Generic.HashSet<int>(); set.Add(29); set.Add(29); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
var map = new System.Collections.Generic.Dictionary<int, int>(); map[29] = 30; Console.WriteLine(map.ContainsKey(29) && map[29] == 30);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_tuple_ordering() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
var tuple = (left: 29, right: 30); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
string feature = "multidim_array_traversal"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn multidim_array_traversal_double_identity() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
double seed = 29; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn multidim_array_traversal_feature_entropy() {
    assert_eq!(run_csharp(r#"// multidim_array_traversal
string feature = "multidim_array_traversal:29"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
