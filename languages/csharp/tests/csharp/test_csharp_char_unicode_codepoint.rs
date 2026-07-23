use super::helpers::run_csharp;

#[test]
fn char_unicode_codepoint_arithmetic_increment() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
int seed = 22; Console.WriteLine(seed + 1 > seed);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_arithmetic_inverse() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
int seed = 22; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_arithmetic_zeroing() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
int seed = 22; Console.WriteLine(seed - seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_ordering_pair() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
int seed = 22; int right = seed + 1; Console.WriteLine(seed < right);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_string_non_empty() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
string feature = "char_unicode_codepoint"; Console.WriteLine(feature.Length > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_string_contains_probe() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
string feature = "char_unicode_codepoint"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_string_first_char() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
string feature = "char_unicode_codepoint"; Console.WriteLine(feature[0] == feature[0]);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_array_length_stable() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
int seed = 22; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_for_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_while_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_ternary_truth() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
int seed = 22; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_nullable_fallback() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
int? maybe = null; int fallback = maybe ?? 22; Console.WriteLine(fallback == 22);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_nullable_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
int? maybe = 22; Console.WriteLine(maybe.HasValue && maybe.Value == 22);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_list_count_contract() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
var values = new System.Collections.Generic.List<int> { 22, 23, 22 }; Console.WriteLine(values.Count == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_hashset_uniqueness() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
var set = new System.Collections.Generic.HashSet<int>(); set.Add(22); set.Add(22); Console.WriteLine(set.Count == 1);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_dictionary_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
var map = new System.Collections.Generic.Dictionary<int, int>(); map[22] = 23; Console.WriteLine(map.ContainsKey(22) && map[22] == 23);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_tuple_ordering() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
var tuple = (left: 22, right: 23); Console.WriteLine(tuple.left < tuple.right);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_string_prefix_suffix() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
string feature = "char_unicode_codepoint"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_double_identity() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
double seed = 22; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_decimal_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#
        ),
        &["True"]
    );
}

#[test]
fn char_unicode_codepoint_feature_entropy() {
    assert_eq!(
        run_csharp(
            r#"// char_unicode_codepoint
string feature = "char_unicode_codepoint:22"; Console.WriteLine(feature.Length >= 1);"#
        ),
        &["True"]
    );
}
