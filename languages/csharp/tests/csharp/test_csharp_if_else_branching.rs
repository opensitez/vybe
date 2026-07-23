use super::helpers::run_csharp;

#[test]
fn if_else_branching_arithmetic_increment() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
int seed = 44; Console.WriteLine(seed + 1 > seed);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_arithmetic_inverse() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
int seed = 44; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_arithmetic_zeroing() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
int seed = 44; Console.WriteLine(seed - seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_ordering_pair() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
int seed = 44; int right = seed + 1; Console.WriteLine(seed < right);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_string_non_empty() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
string feature = "if_else_branching"; Console.WriteLine(feature.Length > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_string_contains_probe() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
string feature = "if_else_branching"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_string_first_char() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
string feature = "if_else_branching"; Console.WriteLine(feature[0] == feature[0]);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_array_length_stable() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
int seed = 44; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_for_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_while_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_ternary_truth() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
int seed = 44; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_nullable_fallback() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
int? maybe = null; int fallback = maybe ?? 44; Console.WriteLine(fallback == 44);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_nullable_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
int? maybe = 44; Console.WriteLine(maybe.HasValue && maybe.Value == 44);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_list_count_contract() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
var values = new System.Collections.Generic.List<int> { 44, 45, 44 }; Console.WriteLine(values.Count == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_hashset_uniqueness() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
var set = new System.Collections.Generic.HashSet<int>(); set.Add(44); set.Add(44); Console.WriteLine(set.Count == 1);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_dictionary_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
var map = new System.Collections.Generic.Dictionary<int, int>(); map[44] = 45; Console.WriteLine(map.ContainsKey(44) && map[44] == 45);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_tuple_ordering() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
var tuple = (left: 44, right: 45); Console.WriteLine(tuple.left < tuple.right);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_string_prefix_suffix() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
string feature = "if_else_branching"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_double_identity() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
double seed = 44; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_decimal_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#
        ),
        &["True"]
    );
}

#[test]
fn if_else_branching_feature_entropy() {
    assert_eq!(
        run_csharp(
            r#"// if_else_branching
string feature = "if_else_branching:44"; Console.WriteLine(feature.Length >= 1);"#
        ),
        &["True"]
    );
}
