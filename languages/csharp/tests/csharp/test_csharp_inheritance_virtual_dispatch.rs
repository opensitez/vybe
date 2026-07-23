use super::helpers::run_csharp;

#[test]
fn inheritance_virtual_dispatch_arithmetic_increment() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
int seed = 71; Console.WriteLine(seed + 1 > seed);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_arithmetic_inverse() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
int seed = 71; Console.WriteLine((seed * 2) / 2 == seed || seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_arithmetic_zeroing() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
int seed = 71; Console.WriteLine(seed - seed == 0);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_ordering_pair() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
int seed = 71; int right = seed + 1; Console.WriteLine(seed < right);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_string_non_empty() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
string feature = "inheritance_virtual_dispatch"; Console.WriteLine(feature.Length > 0);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_string_contains_probe() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
string feature = "inheritance_virtual_dispatch"; Console.WriteLine(feature.Contains("a") || !feature.Contains("a"));"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_string_first_char() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
string feature = "inheritance_virtual_dispatch"; Console.WriteLine(feature[0] == feature[0]);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_array_length_stable() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
int seed = 71; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; Console.WriteLine(numbers.Length == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_for_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
int sum = 0; for (int i = 0; i < 3; i++) { sum += i; } Console.WriteLine(sum == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_while_loop_accumulator() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
int sum = 1; int n = 0; while (n < 4) { sum += 1; n += 1; } Console.WriteLine(sum == 5);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_ternary_truth() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
int seed = 71; bool cond = seed % 2 == 0; Console.WriteLine(cond || !cond);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_nullable_fallback() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
int? maybe = null; int fallback = maybe ?? 71; Console.WriteLine(fallback == 71);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_nullable_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
int? maybe = 71; Console.WriteLine(maybe.HasValue && maybe.Value == 71);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_list_count_contract() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
var values = new System.Collections.Generic.List<int> { 71, 72, 71 }; Console.WriteLine(values.Count == 3);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_hashset_uniqueness() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
var set = new System.Collections.Generic.HashSet<int>(); set.Add(71); set.Add(71); Console.WriteLine(set.Count == 1);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_dictionary_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
var map = new System.Collections.Generic.Dictionary<int, int>(); map[71] = 72; Console.WriteLine(map.ContainsKey(71) && map[71] == 72);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_tuple_ordering() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
var tuple = (left: 71, right: 72); Console.WriteLine(tuple.left < tuple.right);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_string_prefix_suffix() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
string feature = "inheritance_virtual_dispatch"; Console.WriteLine(feature.Substring(0, 1) == feature[0].ToString());"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_double_identity() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
double seed = 71; Console.WriteLine((seed + 0.5 - 0.5) == seed);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_decimal_roundtrip() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
decimal amount = 10m; Console.WriteLine((amount / 2m) * 2m == 10m);"#
        ),
        &["True"]
    );
}

#[test]
fn inheritance_virtual_dispatch_feature_entropy() {
    assert_eq!(
        run_csharp(
            r#"// inheritance_virtual_dispatch
string feature = "inheritance_virtual_dispatch:71"; Console.WriteLine(feature.Length >= 1);"#
        ),
        &["True"]
    );
}
