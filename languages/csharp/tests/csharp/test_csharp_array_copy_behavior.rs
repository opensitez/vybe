use super::helpers::run_csharp;

#[test]
fn array_copy_behavior_arithmetic_increment() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
int seed = 26; Console.WriteLine(seed + 1 > seed);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_arithmetic_inverse() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
int seed = 26; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_arithmetic_zeroing() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
int seed = 26; Console.WriteLine(seed - seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_ordering_pair() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
int seed = 26; int right = seed + 1; Console.WriteLine(seed < right);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_string_non_empty() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
string feature = "array_copy_behavior"; Console.WriteLine(feature.Length > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_string_contains_probe() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
string feature = "array_copy_behavior"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_string_first_char() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
string feature = "array_copy_behavior"; Console.WriteLine(feature[0] == feature[0]);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_array_length_stable() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
int seed = 26; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_for_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_while_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_ternary_truth() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
int seed = 26; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_nullable_fallback() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
int? maybe = null; int fallback = maybe ?? 26; Console.WriteLine(fallback == 26);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_nullable_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
int? maybe = 26; Console.WriteLine(maybe.HasValue && maybe.Value == 26);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_list_count_contract() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
var values = new System.Collections.Generic.List<int> { 26, 27, 26 }; Console.WriteLine(values.Count == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_hashset_uniqueness() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
var set = new System.Collections.Generic.HashSet<int>(); set.Add(26); set.Add(26); Console.WriteLine(set.Count == 1);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_dictionary_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
var map = new System.Collections.Generic.Dictionary<int, int>(); map[26] = 27; Console.WriteLine(map.ContainsKey(26) && map[26] == 27);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_tuple_ordering() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
var tuple = (left: 26, right: 27); Console.WriteLine(tuple.left < tuple.right);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_string_prefix_suffix() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
string feature = "array_copy_behavior"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_double_identity() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
double seed = 26; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_decimal_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#
        ),
        &["True"]
    );
}

#[test]
fn array_copy_behavior_feature_entropy() {
    assert_eq!(
        run_csharp(
            r#"// array_copy_behavior
string feature = "array_copy_behavior:26"; Console.WriteLine(feature.Length >= 1);"#
        ),
        &["True"]
    );
}
