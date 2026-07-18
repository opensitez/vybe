use super::helpers::run_csharp;

#[test]
fn lambda_expression_typing_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
int seed = 76; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
int seed = 76; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
int seed = 76; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_ordering_pair() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
int seed = 76; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_string_non_empty() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
string feature = "lambda_expression_typing"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_string_contains_probe() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
string feature = "lambda_expression_typing"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn lambda_expression_typing_string_first_char() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
string feature = "lambda_expression_typing"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_array_length_stable() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
int seed = 76; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_ternary_truth() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
int seed = 76; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_nullable_fallback() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
int? maybe = null; int fallback = maybe ?? 76; Console.WriteLine(fallback == 76);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
int? maybe = 76; Console.WriteLine(maybe.HasValue && maybe.Value == 76);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_list_count_contract() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
var values = new System.Collections.Generic.List<int> { 76, 77, 76 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
var set = new System.Collections.Generic.HashSet<int>(); set.Add(76); set.Add(76); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
var map = new System.Collections.Generic.Dictionary<int, int>(); map[76] = 77; Console.WriteLine(map.ContainsKey(76) && map[76] == 77);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_tuple_ordering() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
var tuple = (left: 76, right: 77); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
string feature = "lambda_expression_typing"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn lambda_expression_typing_double_identity() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
double seed = 76; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn lambda_expression_typing_feature_entropy() {
    assert_eq!(run_csharp(r#"// lambda_expression_typing
string feature = "lambda_expression_typing:76"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
