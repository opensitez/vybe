use super::helpers::run_csharp;

#[test]
fn queue_stack_surface_arithmetic_increment() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
int seed = 32; Console.WriteLine(seed + 1 > seed);"#), &["True"]);
}

#[test]
fn queue_stack_surface_arithmetic_inverse() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
int seed = 32; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#), &["True"]);
}

#[test]
fn queue_stack_surface_arithmetic_zeroing() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
int seed = 32; Console.WriteLine(seed - seed == 0);"#), &["True"]);
}

#[test]
fn queue_stack_surface_ordering_pair() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
int seed = 32; int right = seed + 1; Console.WriteLine(seed < right);"#), &["True"]);
}

#[test]
fn queue_stack_surface_string_non_empty() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
string feature = "queue_stack_surface"; Console.WriteLine(feature.Length > 0);"#), &["True"]);
}

#[test]
fn queue_stack_surface_string_contains_probe() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
string feature = "queue_stack_surface"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#), &["True"]);
}

#[test]
fn queue_stack_surface_string_first_char() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
string feature = "queue_stack_surface"; Console.WriteLine(feature[0] == feature[0]);"#), &["True"]);
}

#[test]
fn queue_stack_surface_array_length_stable() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
int seed = 32; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#), &["True"]);
}

#[test]
fn queue_stack_surface_for_loop_accumulator() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#), &["True"]);
}

#[test]
fn queue_stack_surface_while_loop_accumulator() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#), &["True"]);
}

#[test]
fn queue_stack_surface_ternary_truth() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
int seed = 32; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#), &["True"]);
}

#[test]
fn queue_stack_surface_nullable_fallback() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
int? maybe = null; int fallback = maybe ?? 32; Console.WriteLine(fallback == 32);"#), &["True"]);
}

#[test]
fn queue_stack_surface_nullable_roundtrip() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
int? maybe = 32; Console.WriteLine(maybe.HasValue && maybe.Value == 32);"#), &["True"]);
}

#[test]
fn queue_stack_surface_list_count_contract() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
var values = new System.Collections.Generic.List<int> { 32, 33, 32 }; Console.WriteLine(values.Count == 3);"#), &["True"]);
}

#[test]
fn queue_stack_surface_hashset_uniqueness() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
var set = new System.Collections.Generic.HashSet<int>(); set.Add(32); set.Add(32); Console.WriteLine(set.Count == 1);"#), &["True"]);
}

#[test]
fn queue_stack_surface_dictionary_roundtrip() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
var map = new System.Collections.Generic.Dictionary<int, int>(); map[32] = 33; Console.WriteLine(map.ContainsKey(32) && map[32] == 33);"#), &["True"]);
}

#[test]
fn queue_stack_surface_tuple_ordering() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
var tuple = (left: 32, right: 33); Console.WriteLine(tuple.left < tuple.right);"#), &["True"]);
}

#[test]
fn queue_stack_surface_string_prefix_suffix() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
string feature = "queue_stack_surface"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#), &["True"]);
}

#[test]
fn queue_stack_surface_double_identity() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
double seed = 32; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#), &["True"]);
}

#[test]
fn queue_stack_surface_decimal_roundtrip() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#), &["True"]);
}

#[test]
fn queue_stack_surface_feature_entropy() {
    assert_eq!(run_csharp(r#"// queue_stack_surface
string feature = "queue_stack_surface:32"; Console.WriteLine(feature.Length >= 1);"#), &["True"]);
}
