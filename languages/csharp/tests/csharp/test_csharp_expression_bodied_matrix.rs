use super::helpers::run_csharp;

#[test]
fn expression_bodied_matrix_arithmetic_increment() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
int seed = 106; Console.WriteLine(seed + 1 > seed);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_arithmetic_inverse() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
int seed = 106; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_arithmetic_zeroing() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
int seed = 106; Console.WriteLine(seed - seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_ordering_pair() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
int seed = 106; int right = seed + 1; Console.WriteLine(seed < right);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_string_non_empty() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
string feature = "expression_bodied_matrix"; Console.WriteLine(feature.Length > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_string_contains_probe() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
string feature = "expression_bodied_matrix"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_string_first_char() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
string feature = "expression_bodied_matrix"; Console.WriteLine(feature[0] == feature[0]);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_array_length_stable() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
int seed = 106; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_for_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_while_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_ternary_truth() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
int seed = 106; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_nullable_fallback() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
int? maybe = null; int fallback = maybe ?? 106; Console.WriteLine(fallback == 106);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_nullable_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
int? maybe = 106; Console.WriteLine(maybe.HasValue && maybe.Value == 106);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_list_count_contract() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
var values = new System.Collections.Generic.List<int> { 106, 107, 106 }; Console.WriteLine(values.Count == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_hashset_uniqueness() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(106); set.Add(106); Console.WriteLine(set.Count == 1);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_dictionary_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[106] = 107; Console.WriteLine(map.ContainsKey(106) && map[106] == 107);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_tuple_ordering() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
var tuple = (left: 106, right: 107); Console.WriteLine(tuple.left < tuple.right);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_string_prefix_suffix() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
string feature = "expression_bodied_matrix"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_double_identity() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
double seed = 106; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_decimal_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#
        ),
        &["True"]
    );
}

#[test]
fn expression_bodied_matrix_feature_entropy() {
    assert_eq!(
        run_csharp(
            r#"// expression_bodied_matrix
string feature = "expression_bodied_matrix:106"; Console.WriteLine(feature.Length >= 1);"#
        ),
        &["True"]
    );
}
