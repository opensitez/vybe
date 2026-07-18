use super::helpers::run_csharp;

#[test]
fn boolean_algebra_surfaces_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
int seed = 11; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
int seed = 11; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
int seed = 11; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_ordering_pair() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
int seed = 11; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_string_non_empty() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
string feature = "boolean_algebra_surfaces"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_string_contains_probe() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
string feature = "boolean_algebra_surfaces"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_string_first_char() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
string feature = "boolean_algebra_surfaces"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_array_length_stable() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
int seed = 11; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_ternary_truth() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
int seed = 11; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_nullable_fallback() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
int? maybe = null; int fallback = maybe ?? 11; Console.WriteLine(fallback == 11);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
int? maybe = 11; Console.WriteLine(maybe.HasValue && maybe.Value == 11);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_list_count_contract() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
var values = new System.Collections.Generic.List<int> { 11, 12, 11 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
var set = new System.Collections.Generic.HashSet<int>(); set.Add(11); set.Add(11); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
var map = new System.Collections.Generic.Dictionary<int, int>(); map[11] = 12; Console.WriteLine(map.ContainsKey(11) && map[11] == 12);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_tuple_ordering() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
var tuple = (left: 11, right: 12); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
string feature = "boolean_algebra_surfaces"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_double_identity() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
double seed = 11; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn boolean_algebra_surfaces_feature_entropy() {
    assert_eq!(run_csharp(r#"// boolean_algebra_surfaces
string feature = "boolean_algebra_surfaces:11"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
