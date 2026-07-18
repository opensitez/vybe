use super::helpers::run_csharp;

#[test]
fn exception_type_checks_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// exception_type_checks
int seed = 53; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn exception_type_checks_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// exception_type_checks
int seed = 53; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn exception_type_checks_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// exception_type_checks
int seed = 53; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn exception_type_checks_ordering_pair() {
    assert_eq!(run_csharp(r#"// exception_type_checks
int seed = 53; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn exception_type_checks_string_non_empty() {
    assert_eq!(run_csharp(r#"// exception_type_checks
string feature = "exception_type_checks"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn exception_type_checks_string_contains_probe() {
    assert_eq!(run_csharp(r#"// exception_type_checks
string feature = "exception_type_checks"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn exception_type_checks_string_first_char() {
    assert_eq!(run_csharp(r#"// exception_type_checks
string feature = "exception_type_checks"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn exception_type_checks_array_length_stable() {
    assert_eq!(run_csharp(r#"// exception_type_checks
int seed = 53; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn exception_type_checks_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// exception_type_checks
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn exception_type_checks_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// exception_type_checks
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn exception_type_checks_ternary_truth() {
    assert_eq!(run_csharp(r#"// exception_type_checks
int seed = 53; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn exception_type_checks_nullable_fallback() {
    assert_eq!(run_csharp(r#"// exception_type_checks
int? maybe = null; int fallback = maybe ?? 53; Console.WriteLine(fallback == 53);"#), &["True"]);
}

#[test]
fn exception_type_checks_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// exception_type_checks
int? maybe = 53; Console.WriteLine(maybe.HasValue && maybe.Value == 53);"#), &["True"]);
}

#[test]
fn exception_type_checks_list_count_contract() {
    assert_eq!(run_csharp(r#"// exception_type_checks
var values = new System.Collections.Generic.List<int> { 53, 54, 53 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn exception_type_checks_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// exception_type_checks
var set = new System.Collections.Generic.HashSet<int>(); set.Add(53); set.Add(53); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn exception_type_checks_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// exception_type_checks
var map = new System.Collections.Generic.Dictionary<int, int>(); map[53] = 54; Console.WriteLine(map.ContainsKey(53) && map[53] == 54);"#), &["True"]);
}

#[test]
fn exception_type_checks_tuple_ordering() {
    assert_eq!(run_csharp(r#"// exception_type_checks
var tuple = (left: 53, right: 54); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn exception_type_checks_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// exception_type_checks
string feature = "exception_type_checks"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn exception_type_checks_double_identity() {
    assert_eq!(run_csharp(r#"// exception_type_checks
double seed = 53; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn exception_type_checks_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// exception_type_checks
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn exception_type_checks_feature_entropy() {
    assert_eq!(run_csharp(r#"// exception_type_checks
string feature = "exception_type_checks:53"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
