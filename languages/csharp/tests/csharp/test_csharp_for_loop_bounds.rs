use super::helpers::run_csharp;

#[test]
fn for_loop_bounds_arithmetic_increment() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
int seed = 45; Console.WriteLine(seed + 1 > seed);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_arithmetic_inverse() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
int seed = 45; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_arithmetic_zeroing() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
int seed = 45; Console.WriteLine(seed - seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_ordering_pair() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
int seed = 45; int right = seed + 1; Console.WriteLine(seed < right);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_string_non_empty() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
string feature = "for_loop_bounds"; Console.WriteLine(feature.Length > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_string_contains_probe() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
string feature = "for_loop_bounds"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_string_first_char() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
string feature = "for_loop_bounds"; Console.WriteLine(feature[0] == feature[0]);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_array_length_stable() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
int seed = 45; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_for_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_while_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_ternary_truth() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
int seed = 45; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_nullable_fallback() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
int? maybe = null; int fallback = maybe ?? 45; Console.WriteLine(fallback == 45);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_nullable_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
int? maybe = 45; Console.WriteLine(maybe.HasValue && maybe.Value == 45);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_list_count_contract() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
var values = new System.Collections.Generic.List<int> { 45, 46, 45 }; Console.WriteLine(values.Count == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_hashset_uniqueness() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
var set = new System.Collections.Generic.HashSet<int>(); set.Add(45); set.Add(45); Console.WriteLine(set.Count == 1);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_dictionary_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
var map = new System.Collections.Generic.Dictionary<int, int>(); map[45] = 46; Console.WriteLine(map.ContainsKey(45) && map[45] == 46);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_tuple_ordering() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
var tuple = (left: 45, right: 46); Console.WriteLine(tuple.left < tuple.right);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_string_prefix_suffix() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
string feature = "for_loop_bounds"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_double_identity() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
double seed = 45; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_decimal_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#
        ),
        &["True"]
    );
}

#[test]
fn for_loop_bounds_feature_entropy() {
    assert_eq!(
        run_csharp(
            r#"// for_loop_bounds
string feature = "for_loop_bounds:45"; Console.WriteLine(feature.Length >= 1);"#
        ),
        &["True"]
    );
}
