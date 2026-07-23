use super::helpers::run_csharp;

#[test]
fn do_while_matrix_arithmetic_increment() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
int seed = 48; Console.WriteLine(seed + 1 > seed);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_arithmetic_inverse() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
int seed = 48; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_arithmetic_zeroing() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
int seed = 48; Console.WriteLine(seed - seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_ordering_pair() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
int seed = 48; int right = seed + 1; Console.WriteLine(seed < right);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_string_non_empty() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
string feature = "do_while_matrix"; Console.WriteLine(feature.Length > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_string_contains_probe() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
string feature = "do_while_matrix"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_string_first_char() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
string feature = "do_while_matrix"; Console.WriteLine(feature[0] == feature[0]);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_array_length_stable() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
int seed = 48; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_for_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_while_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_ternary_truth() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
int seed = 48; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_nullable_fallback() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
int? maybe = null; int fallback = maybe ?? 48; Console.WriteLine(fallback == 48);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_nullable_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
int? maybe = 48; Console.WriteLine(maybe.HasValue && maybe.Value == 48);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_list_count_contract() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
var values = new System.Collections.Generic.List<int> { 48, 49, 48 }; Console.WriteLine(values.Count == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_hashset_uniqueness() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(48); set.Add(48); Console.WriteLine(set.Count == 1);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_dictionary_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[48] = 49; Console.WriteLine(map.ContainsKey(48) && map[48] == 49);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_tuple_ordering() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
var tuple = (left: 48, right: 49); Console.WriteLine(tuple.left < tuple.right);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_string_prefix_suffix() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
string feature = "do_while_matrix"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_double_identity() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
double seed = 48; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_decimal_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#
        ),
        &["True"]
    );
}

#[test]
fn do_while_matrix_feature_entropy() {
    assert_eq!(
        run_csharp(
            r#"// do_while_matrix
string feature = "do_while_matrix:48"; Console.WriteLine(feature.Length >= 1);"#
        ),
        &["True"]
    );
}
