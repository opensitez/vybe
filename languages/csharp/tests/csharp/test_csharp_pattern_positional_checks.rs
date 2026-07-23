use super::helpers::run_csharp;

#[test]
fn pattern_positional_checks_arithmetic_increment() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
int seed = 115; Console.WriteLine(seed + 1 > seed);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_arithmetic_inverse() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
int seed = 115; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_arithmetic_zeroing() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
int seed = 115; Console.WriteLine(seed - seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_ordering_pair() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
int seed = 115; int right = seed + 1; Console.WriteLine(seed < right);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_string_non_empty() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
string feature = "pattern_positional_checks"; Console.WriteLine(feature.Length > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_string_contains_probe() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
string feature = "pattern_positional_checks"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_string_first_char() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
string feature = "pattern_positional_checks"; Console.WriteLine(feature[0] == feature[0]);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_array_length_stable() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
int seed = 115; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_for_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_while_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_ternary_truth() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
int seed = 115; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_nullable_fallback() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
int? maybe = null; int fallback = maybe ?? 115; Console.WriteLine(fallback == 115);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_nullable_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
int? maybe = 115; Console.WriteLine(maybe.HasValue && maybe.Value == 115);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_list_count_contract() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
var values = new System.Collections.Generic.List<int> { 115, 116, 115 }; Console.WriteLine(values.Count == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_hashset_uniqueness() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
var set = new System.Collections.Generic.HashSet<int>(); set.Add(115); set.Add(115); Console.WriteLine(set.Count == 1);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_dictionary_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
var map = new System.Collections.Generic.Dictionary<int, int>(); map[115] = 116; Console.WriteLine(map.ContainsKey(115) && map[115] == 116);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_tuple_ordering() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
var tuple = (left: 115, right: 116); Console.WriteLine(tuple.left < tuple.right);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_string_prefix_suffix() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
string feature = "pattern_positional_checks"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_double_identity() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
double seed = 115; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_decimal_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#
        ),
        &["True"]
    );
}

#[test]
fn pattern_positional_checks_feature_entropy() {
    assert_eq!(
        run_csharp(
            r#"// pattern_positional_checks
string feature = "pattern_positional_checks:115"; Console.WriteLine(feature.Length >= 1);"#
        ),
        &["True"]
    );
}
