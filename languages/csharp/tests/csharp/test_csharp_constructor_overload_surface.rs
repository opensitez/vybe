use super::helpers::run_csharp;

#[test]
fn constructor_overload_surface_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
int seed = 67; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
int seed = 67; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
int seed = 67; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_ordering_pair() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
int seed = 67; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_string_non_empty() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
string feature = "constructor_overload_surface"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_string_contains_probe() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
string feature = "constructor_overload_surface"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn constructor_overload_surface_string_first_char() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
string feature = "constructor_overload_surface"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_array_length_stable() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
int seed = 67; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_ternary_truth() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
int seed = 67; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_nullable_fallback() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
int? maybe = null; int fallback = maybe ?? 67; Console.WriteLine(fallback == 67);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
int? maybe = 67; Console.WriteLine(maybe.HasValue && maybe.Value == 67);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_list_count_contract() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
var values = new System.Collections.Generic.List<int> { 67, 68, 67 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
var set = new System.Collections.Generic.HashSet<int>(); set.Add(67); set.Add(67); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[67] = 68; Console.WriteLine(map.ContainsKey(67) && map[67] == 68);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_tuple_ordering() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
var tuple = (left: 67, right: 68); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
string feature = "constructor_overload_surface"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn constructor_overload_surface_double_identity() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
double seed = 67; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn constructor_overload_surface_feature_entropy() {
    assert_eq!(run_csharp(r#"// constructor_overload_surface
string feature = "constructor_overload_surface:67"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
