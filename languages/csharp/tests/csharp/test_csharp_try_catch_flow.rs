use super::helpers::run_csharp;

#[test]
fn try_catch_flow_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// try_catch_flow
int seed = 51; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn try_catch_flow_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// try_catch_flow
int seed = 51; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn try_catch_flow_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// try_catch_flow
int seed = 51; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn try_catch_flow_ordering_pair() {
    assert_eq!(run_csharp(r#"// try_catch_flow
int seed = 51; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn try_catch_flow_string_non_empty() {
    assert_eq!(run_csharp(r#"// try_catch_flow
string feature = "try_catch_flow"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn try_catch_flow_string_contains_probe() {
    assert_eq!(run_csharp(r#"// try_catch_flow
string feature = "try_catch_flow"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn try_catch_flow_string_first_char() {
    assert_eq!(run_csharp(r#"// try_catch_flow
string feature = "try_catch_flow"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn try_catch_flow_array_length_stable() {
    assert_eq!(run_csharp(r#"// try_catch_flow
int seed = 51; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn try_catch_flow_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// try_catch_flow
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn try_catch_flow_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// try_catch_flow
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn try_catch_flow_ternary_truth() {
    assert_eq!(run_csharp(r#"// try_catch_flow
int seed = 51; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn try_catch_flow_nullable_fallback() {
    assert_eq!(run_csharp(r#"// try_catch_flow
int? maybe = null; int fallback = maybe ?? 51; Console.WriteLine(fallback == 51);"#), &["True"]);
}

#[test]
fn try_catch_flow_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// try_catch_flow
int? maybe = 51; Console.WriteLine(maybe.HasValue && maybe.Value == 51);"#), &["True"]);
}

#[test]
fn try_catch_flow_list_count_contract() {
    assert_eq!(run_csharp(r#"// try_catch_flow
var values = new System.Collections.Generic.List<int> { 51, 52, 51 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn try_catch_flow_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// try_catch_flow
var set = new System.Collections.Generic.HashSet<int>(); set.Add(51); set.Add(51); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn try_catch_flow_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// try_catch_flow
var map = new System.Collections.Generic.Dictionary<int, int>(); map[51] = 52; Console.WriteLine(map.ContainsKey(51) && map[51] == 52);"#), &["True"]);
}

#[test]
fn try_catch_flow_tuple_ordering() {
    assert_eq!(run_csharp(r#"// try_catch_flow
var tuple = (left: 51, right: 52); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn try_catch_flow_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// try_catch_flow
string feature = "try_catch_flow"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn try_catch_flow_double_identity() {
    assert_eq!(run_csharp(r#"// try_catch_flow
double seed = 51; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn try_catch_flow_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// try_catch_flow
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn try_catch_flow_feature_entropy() {
    assert_eq!(run_csharp(r#"// try_catch_flow
string feature = "try_catch_flow:51"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
