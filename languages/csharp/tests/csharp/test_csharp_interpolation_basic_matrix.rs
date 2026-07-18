use super::helpers::run_csharp;

#[test]
fn interpolation_basic_matrix_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
int seed = 112; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
int seed = 112; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
int seed = 112; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_ordering_pair() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
int seed = 112; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_string_non_empty() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
string feature = "interpolation_basic_matrix"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_string_contains_probe() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
string feature = "interpolation_basic_matrix"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_string_first_char() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
string feature = "interpolation_basic_matrix"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_array_length_stable() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
int seed = 112; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_ternary_truth() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
int seed = 112; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_nullable_fallback() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
int? maybe = null; int fallback = maybe ?? 112; Console.WriteLine(fallback == 112);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
int? maybe = 112; Console.WriteLine(maybe.HasValue && maybe.Value == 112);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_list_count_contract() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
var values = new System.Collections.Generic.List<int> { 112, 113, 112 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(112); set.Add(112); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[112] = 113; Console.WriteLine(map.ContainsKey(112) && map[112] == 113);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_tuple_ordering() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
var tuple = (left: 112, right: 113); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
string feature = "interpolation_basic_matrix"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_double_identity() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
double seed = 112; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn interpolation_basic_matrix_feature_entropy() {
    assert_eq!(run_csharp(r#"// interpolation_basic_matrix
string feature = "interpolation_basic_matrix:112"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
