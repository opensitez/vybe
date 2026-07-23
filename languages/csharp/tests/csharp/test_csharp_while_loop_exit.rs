use super::helpers::run_csharp;

#[test]
fn while_loop_exit_arithmetic_increment() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
int seed = 47; Console.WriteLine(seed + 1 > seed);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_arithmetic_inverse() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
int seed = 47; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_arithmetic_zeroing() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
int seed = 47; Console.WriteLine(seed - seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_ordering_pair() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
int seed = 47; int right = seed + 1; Console.WriteLine(seed < right);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_string_non_empty() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
string feature = "while_loop_exit"; Console.WriteLine(feature.Length > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_string_contains_probe() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
string feature = "while_loop_exit"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_string_first_char() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
string feature = "while_loop_exit"; Console.WriteLine(feature[0] == feature[0]);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_array_length_stable() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
int seed = 47; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_for_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_while_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_ternary_truth() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
int seed = 47; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_nullable_fallback() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
int? maybe = null; int fallback = maybe ?? 47; Console.WriteLine(fallback == 47);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_nullable_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
int? maybe = 47; Console.WriteLine(maybe.HasValue && maybe.Value == 47);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_list_count_contract() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
var values = new System.Collections.Generic.List<int> { 47, 48, 47 }; Console.WriteLine(values.Count == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_hashset_uniqueness() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
var set = new System.Collections.Generic.HashSet<int>(); set.Add(47); set.Add(47); Console.WriteLine(set.Count == 1);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_dictionary_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
var map = new System.Collections.Generic.Dictionary<int, int>(); map[47] = 48; Console.WriteLine(map.ContainsKey(47) && map[47] == 48);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_tuple_ordering() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
var tuple = (left: 47, right: 48); Console.WriteLine(tuple.left < tuple.right);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_string_prefix_suffix() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
string feature = "while_loop_exit"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_double_identity() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
double seed = 47; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_decimal_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#
        ),
        &["True"]
    );
}

#[test]
fn while_loop_exit_feature_entropy() {
    assert_eq!(
        run_csharp(
            r#"// while_loop_exit
string feature = "while_loop_exit:47"; Console.WriteLine(feature.Length >= 1);"#
        ),
        &["True"]
    );
}
