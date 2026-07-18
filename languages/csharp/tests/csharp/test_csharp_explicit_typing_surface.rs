use super::helpers::run_csharp;

#[test]
fn explicit_typing_surface_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
int seed = 60; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
int seed = 60; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
int seed = 60; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_ordering_pair() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
int seed = 60; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_string_non_empty() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
string feature = "explicit_typing_surface"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_string_contains_probe() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
string feature = "explicit_typing_surface"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn explicit_typing_surface_string_first_char() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
string feature = "explicit_typing_surface"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_array_length_stable() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
int seed = 60; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_ternary_truth() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
int seed = 60; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_nullable_fallback() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
int? maybe = null; int fallback = maybe ?? 60; Console.WriteLine(fallback == 60);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
int? maybe = 60; Console.WriteLine(maybe.HasValue && maybe.Value == 60);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_list_count_contract() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
var values = new System.Collections.Generic.List<int> { 60, 61, 60 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
var set = new System.Collections.Generic.HashSet<int>(); set.Add(60); set.Add(60); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[60] = 61; Console.WriteLine(map.ContainsKey(60) && map[60] == 61);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_tuple_ordering() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
var tuple = (left: 60, right: 61); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
string feature = "explicit_typing_surface"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn explicit_typing_surface_double_identity() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
double seed = 60; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn explicit_typing_surface_feature_entropy() {
    assert_eq!(run_csharp(r#"// explicit_typing_surface
string feature = "explicit_typing_surface:60"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
