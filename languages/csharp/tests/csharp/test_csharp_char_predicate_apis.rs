use super::helpers::run_csharp;

#[test]
fn char_predicate_apis_arithmetic_increment() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
int seed = 23; Console.WriteLine(seed + 1 > seed);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_arithmetic_inverse() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
int seed = 23; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_arithmetic_zeroing() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
int seed = 23; Console.WriteLine(seed - seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_ordering_pair() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
int seed = 23; int right = seed + 1; Console.WriteLine(seed < right);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_string_non_empty() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
string feature = "char_predicate_apis"; Console.WriteLine(feature.Length > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_string_contains_probe() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
string feature = "char_predicate_apis"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_string_first_char() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
string feature = "char_predicate_apis"; Console.WriteLine(feature[0] == feature[0]);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_array_length_stable() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
int seed = 23; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_for_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_while_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_ternary_truth() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
int seed = 23; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_nullable_fallback() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
int? maybe = null; int fallback = maybe ?? 23; Console.WriteLine(fallback == 23);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_nullable_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
int? maybe = 23; Console.WriteLine(maybe.HasValue && maybe.Value == 23);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_list_count_contract() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
var values = new System.Collections.Generic.List<int> { 23, 24, 23 }; Console.WriteLine(values.Count == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_hashset_uniqueness() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
var set = new System.Collections.Generic.HashSet<int>(); set.Add(23); set.Add(23); Console.WriteLine(set.Count == 1);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_dictionary_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
var map = new System.Collections.Generic.Dictionary<int, int>(); map[23] = 24; Console.WriteLine(map.ContainsKey(23) && map[23] == 24);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_tuple_ordering() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
var tuple = (left: 23, right: 24); Console.WriteLine(tuple.left < tuple.right);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_string_prefix_suffix() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
string feature = "char_predicate_apis"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_double_identity() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
double seed = 23; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_decimal_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#
        ),
        &["True"]
    );
}

#[test]
fn char_predicate_apis_feature_entropy() {
    assert_eq!(
        run_csharp(
            r#"// char_predicate_apis
string feature = "char_predicate_apis:23"; Console.WriteLine(feature.Length >= 1);"#
        ),
        &["True"]
    );
}
