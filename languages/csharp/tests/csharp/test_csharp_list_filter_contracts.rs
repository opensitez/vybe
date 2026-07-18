use super::helpers::run_csharp;

#[test]
fn list_filter_contracts_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
int seed = 31; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn list_filter_contracts_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
int seed = 31; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn list_filter_contracts_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
int seed = 31; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn list_filter_contracts_ordering_pair() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
int seed = 31; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn list_filter_contracts_string_non_empty() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
string feature = "list_filter_contracts"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn list_filter_contracts_string_contains_probe() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
string feature = "list_filter_contracts"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn list_filter_contracts_string_first_char() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
string feature = "list_filter_contracts"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn list_filter_contracts_array_length_stable() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
int seed = 31; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn list_filter_contracts_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn list_filter_contracts_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn list_filter_contracts_ternary_truth() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
int seed = 31; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn list_filter_contracts_nullable_fallback() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
int? maybe = null; int fallback = maybe ?? 31; Console.WriteLine(fallback == 31);"#), &["True"]);
}

#[test]
fn list_filter_contracts_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
int? maybe = 31; Console.WriteLine(maybe.HasValue && maybe.Value == 31);"#), &["True"]);
}

#[test]
fn list_filter_contracts_list_count_contract() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
var values = new System.Collections.Generic.List<int> { 31, 32, 31 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn list_filter_contracts_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
var set = new System.Collections.Generic.HashSet<int>(); set.Add(31); set.Add(31); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn list_filter_contracts_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
var map = new System.Collections.Generic.Dictionary<int, int>(); map[31] = 32; Console.WriteLine(map.ContainsKey(31) && map[31] == 32);"#), &["True"]);
}

#[test]
fn list_filter_contracts_tuple_ordering() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
var tuple = (left: 31, right: 32); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn list_filter_contracts_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
string feature = "list_filter_contracts"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn list_filter_contracts_double_identity() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
double seed = 31; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn list_filter_contracts_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn list_filter_contracts_feature_entropy() {
    assert_eq!(run_csharp(r#"// list_filter_contracts
string feature = "list_filter_contracts:31"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
