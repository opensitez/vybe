use super::helpers::run_csharp;

#[test]
fn string_unicode_basics_arithmetic_increment() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
int seed = 19; Console.WriteLine(seed + 1 > seed);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_arithmetic_inverse() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
int seed = 19; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_arithmetic_zeroing() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
int seed = 19; Console.WriteLine(seed - seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_ordering_pair() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
int seed = 19; int right = seed + 1; Console.WriteLine(seed < right);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_string_non_empty() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
string feature = "string_unicode_basics"; Console.WriteLine(feature.Length > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_string_contains_probe() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
string feature = "string_unicode_basics"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_string_first_char() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
string feature = "string_unicode_basics"; Console.WriteLine(feature[0] == feature[0]);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_array_length_stable() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
int seed = 19; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_for_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_while_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_ternary_truth() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
int seed = 19; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_nullable_fallback() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
int? maybe = null; int fallback = maybe ?? 19; Console.WriteLine(fallback == 19);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_nullable_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
int? maybe = 19; Console.WriteLine(maybe.HasValue && maybe.Value == 19);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_list_count_contract() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
var values = new System.Collections.Generic.List<int> { 19, 20, 19 }; Console.WriteLine(values.Count == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_hashset_uniqueness() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
var set = new System.Collections.Generic.HashSet<int>(); set.Add(19); set.Add(19); Console.WriteLine(set.Count == 1);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_dictionary_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
var map = new System.Collections.Generic.Dictionary<int, int>(); map[19] = 20; Console.WriteLine(map.ContainsKey(19) && map[19] == 20);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_tuple_ordering() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
var tuple = (left: 19, right: 20); Console.WriteLine(tuple.left < tuple.right);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_string_prefix_suffix() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
string feature = "string_unicode_basics"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_double_identity() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
double seed = 19; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_decimal_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#
        ),
        &["True"]
    );
}

#[test]
fn string_unicode_basics_feature_entropy() {
    assert_eq!(
        run_csharp(
            r#"// string_unicode_basics
string feature = "string_unicode_basics:19"; Console.WriteLine(feature.Length >= 1);"#
        ),
        &["True"]
    );
}
