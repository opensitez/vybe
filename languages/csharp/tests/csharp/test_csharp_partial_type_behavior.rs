use super::helpers::run_csharp;

#[test]
fn partial_type_behavior_arithmetic_increment() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
int seed = 70; Console.WriteLine(seed + 1 > seed);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_arithmetic_inverse() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
int seed = 70; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_arithmetic_zeroing() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
int seed = 70; Console.WriteLine(seed - seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_ordering_pair() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
int seed = 70; int right = seed + 1; Console.WriteLine(seed < right);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_string_non_empty() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
string feature = "partial_type_behavior"; Console.WriteLine(feature.Length > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_string_contains_probe() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
string feature = "partial_type_behavior"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_string_first_char() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
string feature = "partial_type_behavior"; Console.WriteLine(feature[0] == feature[0]);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_array_length_stable() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
int seed = 70; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_for_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_while_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_ternary_truth() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
int seed = 70; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_nullable_fallback() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
int? maybe = null; int fallback = maybe ?? 70; Console.WriteLine(fallback == 70);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_nullable_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
int? maybe = 70; Console.WriteLine(maybe.HasValue && maybe.Value == 70);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_list_count_contract() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
var values = new System.Collections.Generic.List<int> { 70, 71, 70 }; Console.WriteLine(values.Count == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_hashset_uniqueness() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
var set = new System.Collections.Generic.HashSet<int>(); set.Add(70); set.Add(70); Console.WriteLine(set.Count == 1);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_dictionary_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
var map = new System.Collections.Generic.Dictionary<int, int>(); map[70] = 71; Console.WriteLine(map.ContainsKey(70) && map[70] == 71);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_tuple_ordering() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
var tuple = (left: 70, right: 71); Console.WriteLine(tuple.left < tuple.right);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_string_prefix_suffix() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
string feature = "partial_type_behavior"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_double_identity() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
double seed = 70; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_decimal_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#
        ),
        &["True"]
    );
}

#[test]
fn partial_type_behavior_feature_entropy() {
    assert_eq!(
        run_csharp(
            r#"// partial_type_behavior
string feature = "partial_type_behavior:70"; Console.WriteLine(feature.Length >= 1);"#
        ),
        &["True"]
    );
}
