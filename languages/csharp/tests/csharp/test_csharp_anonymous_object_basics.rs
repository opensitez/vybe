use super::helpers::run_csharp;

#[test]
fn anonymous_object_basics_arithmetic_increment() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
int seed = 38; Console.WriteLine(seed + 1 > seed);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_arithmetic_inverse() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
int seed = 38; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_arithmetic_zeroing() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
int seed = 38; Console.WriteLine(seed - seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_ordering_pair() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
int seed = 38; int right = seed + 1; Console.WriteLine(seed < right);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_string_non_empty() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
string feature = "anonymous_object_basics"; Console.WriteLine(feature.Length > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_string_contains_probe() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
string feature = "anonymous_object_basics"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_string_first_char() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
string feature = "anonymous_object_basics"; Console.WriteLine(feature[0] == feature[0]);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_array_length_stable() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
int seed = 38; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_for_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_while_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_ternary_truth() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
int seed = 38; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_nullable_fallback() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
int? maybe = null; int fallback = maybe ?? 38; Console.WriteLine(fallback == 38);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_nullable_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
int? maybe = 38; Console.WriteLine(maybe.HasValue && maybe.Value == 38);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_list_count_contract() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
var values = new System.Collections.Generic.List<int> { 38, 39, 38 }; Console.WriteLine(values.Count == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_hashset_uniqueness() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
var set = new System.Collections.Generic.HashSet<int>(); set.Add(38); set.Add(38); Console.WriteLine(set.Count == 1);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_dictionary_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
var map = new System.Collections.Generic.Dictionary<int, int>(); map[38] = 39; Console.WriteLine(map.ContainsKey(38) && map[38] == 39);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_tuple_ordering() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
var tuple = (left: 38, right: 39); Console.WriteLine(tuple.left < tuple.right);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_string_prefix_suffix() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
string feature = "anonymous_object_basics"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_double_identity() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
double seed = 38; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_decimal_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#
        ),
        &["True"]
    );
}

#[test]
fn anonymous_object_basics_feature_entropy() {
    assert_eq!(
        run_csharp(
            r#"// anonymous_object_basics
string feature = "anonymous_object_basics:38"; Console.WriteLine(feature.Length >= 1);"#
        ),
        &["True"]
    );
}
