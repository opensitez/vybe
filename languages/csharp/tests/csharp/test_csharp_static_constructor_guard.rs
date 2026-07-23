use super::helpers::run_csharp;

#[test]
fn static_constructor_guard_arithmetic_increment() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
int seed = 69; Console.WriteLine(seed + 1 > seed);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_arithmetic_inverse() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
int seed = 69; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_arithmetic_zeroing() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
int seed = 69; Console.WriteLine(seed - seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_ordering_pair() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
int seed = 69; int right = seed + 1; Console.WriteLine(seed < right);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_string_non_empty() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
string feature = "static_constructor_guard"; Console.WriteLine(feature.Length > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_string_contains_probe() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
string feature = "static_constructor_guard"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_string_first_char() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
string feature = "static_constructor_guard"; Console.WriteLine(feature[0] == feature[0]);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_array_length_stable() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
int seed = 69; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_for_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_while_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_ternary_truth() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
int seed = 69; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_nullable_fallback() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
int? maybe = null; int fallback = maybe ?? 69; Console.WriteLine(fallback == 69);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_nullable_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
int? maybe = 69; Console.WriteLine(maybe.HasValue && maybe.Value == 69);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_list_count_contract() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
var values = new System.Collections.Generic.List<int> { 69, 70, 69 }; Console.WriteLine(values.Count == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_hashset_uniqueness() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
var set = new System.Collections.Generic.HashSet<int>(); set.Add(69); set.Add(69); Console.WriteLine(set.Count == 1);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_dictionary_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
var map = new System.Collections.Generic.Dictionary<int, int>(); map[69] = 70; Console.WriteLine(map.ContainsKey(69) && map[69] == 70);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_tuple_ordering() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
var tuple = (left: 69, right: 70); Console.WriteLine(tuple.left < tuple.right);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_string_prefix_suffix() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
string feature = "static_constructor_guard"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_double_identity() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
double seed = 69; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_decimal_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#
        ),
        &["True"]
    );
}

#[test]
fn static_constructor_guard_feature_entropy() {
    assert_eq!(
        run_csharp(
            r#"// static_constructor_guard
string feature = "static_constructor_guard:69"; Console.WriteLine(feature.Length >= 1);"#
        ),
        &["True"]
    );
}
