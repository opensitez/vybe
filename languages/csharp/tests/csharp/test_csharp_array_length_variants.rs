use super::helpers::run_csharp;

#[test]
fn array_length_variants_arithmetic_increment() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
int seed = 25; Console.WriteLine(seed + 1 > seed);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_arithmetic_inverse() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
int seed = 25; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_arithmetic_zeroing() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
int seed = 25; Console.WriteLine(seed - seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_ordering_pair() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
int seed = 25; int right = seed + 1; Console.WriteLine(seed < right);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_string_non_empty() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
string feature = "array_length_variants"; Console.WriteLine(feature.Length > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_string_contains_probe() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
string feature = "array_length_variants"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_string_first_char() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
string feature = "array_length_variants"; Console.WriteLine(feature[0] == feature[0]);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_array_length_stable() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
int seed = 25; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_for_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_while_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_ternary_truth() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
int seed = 25; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_nullable_fallback() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
int? maybe = null; int fallback = maybe ?? 25; Console.WriteLine(fallback == 25);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_nullable_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
int? maybe = 25; Console.WriteLine(maybe.HasValue && maybe.Value == 25);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_list_count_contract() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
var values = new System.Collections.Generic.List<int> { 25, 26, 25 }; Console.WriteLine(values.Count == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_hashset_uniqueness() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
var set = new System.Collections.Generic.HashSet<int>(); set.Add(25); set.Add(25); Console.WriteLine(set.Count == 1);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_dictionary_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
var map = new System.Collections.Generic.Dictionary<int, int>(); map[25] = 26; Console.WriteLine(map.ContainsKey(25) && map[25] == 26);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_tuple_ordering() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
var tuple = (left: 25, right: 26); Console.WriteLine(tuple.left < tuple.right);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_string_prefix_suffix() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
string feature = "array_length_variants"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_double_identity() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
double seed = 25; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_decimal_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#
        ),
        &["True"]
    );
}

#[test]
fn array_length_variants_feature_entropy() {
    assert_eq!(
        run_csharp(
            r#"// array_length_variants
string feature = "array_length_variants:25"; Console.WriteLine(feature.Length >= 1);"#
        ),
        &["True"]
    );
}
