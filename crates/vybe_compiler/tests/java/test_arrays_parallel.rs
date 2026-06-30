use crate::helpers::run_main;

#[test]
fn arrays_parallel_sort_orders_ints() {
    let out = run_main(
        "int[] a = {3, 1, 2}; java.util.Arrays.parallelSort(a); System.out.println(a[0]); System.out.println(a[2]);",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn arrays_parallel_sort_with_comparator() {
    let out = run_main(
        "String[] a = {\"c\", \"a\", \"b\"}; java.util.Arrays.parallelSort(a); System.out.println(a[0]);",
    );
    assert_eq!(out, vec!["a"]);
}

#[test]
fn arrays_parallel_prefix_sum() {
    let out = run_main(
        "int[] a = {1, 2, 3}; java.util.Arrays.parallelPrefix(a, (x, y) -> x + y); System.out.println(a[2]);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn arrays_parallel_prefix_with_identity() {
    let out = run_main(
        "int[] a = {1, 2, 3}; java.util.Arrays.parallelPrefix(0, a, (x, y) -> x + y); System.out.println(a[2]);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn arrays_compare_different_lengths() {
    let out = run_main(
        "int[] a = {1, 2}; int[] b = {1, 2, 3}; System.out.println(java.util.Arrays.compare(a, b) < 0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arrays_compare_identical() {
    let out = run_main(
        "int[] a = {1, 2}; int[] b = {1, 2}; System.out.println(java.util.Arrays.compare(a, b));",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn arrays_compare_less() {
    let out = run_main(
        "int[] a = {1, 2}; int[] b = {1, 3}; System.out.println(java.util.Arrays.compare(a, b) < 0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arrays_compare_greater() {
    let out = run_main(
        "int[] a = {2, 0}; int[] b = {1, 9}; System.out.println(java.util.Arrays.compare(a, b) > 0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arrays_equals_deep_nested() {
    let out = run_main(
        "int[][] a = {{1, 2}, {3}}; int[][] b = {{1, 2}, {3}}; System.out.println(java.util.Arrays.deepEquals(a, b));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arrays_deep_to_string() {
    let out = run_main(
        "int[][] a = {{1, 2}, {3}}; String s = java.util.Arrays.deepToString(a); System.out.println(s.contains(\"1\"));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arrays_set_all_fill() {
    let out = run_main(
        "int[] a = new int[3]; java.util.Arrays.setAll(a, i -> i + 1); System.out.println(a[0]); System.out.println(a[2]);",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn arrays_parallel_set_all() {
    let out = run_main(
        "int[] a = new int[4]; java.util.Arrays.parallelSetAll(a, i -> i * 2); System.out.println(a[3]);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn arrays_fill_parallel() {
    let out = run_main(
        "int[] a = new int[3]; java.util.Arrays.parallelSetAll(a, i -> 7); System.out.println(a[1]);",
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn arrays_stream_longs_sum() {
    let out = run_main(
        "long s = java.util.Arrays.stream(new long[]{1L, 2L, 3L}).sum(); System.out.println(s);",
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn arrays_stream_doubles_average() {
    let out = run_main(
        "double a = java.util.Arrays.stream(new double[]{2.0, 4.0}).average().getAsDouble(); System.out.println((int) a);",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn arrays_copy_of_range() {
    let out = run_main(
        "int[] a = {1, 2, 3, 4}; int[] b = java.util.Arrays.copyOfRange(a, 1, 3); System.out.println(b.length); System.out.println(b[1]);",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn arrays_sort_partial_range() {
    let out = run_main(
        "int[] a = {3, 2, 1, 0}; java.util.Arrays.sort(a, 0, 3); System.out.println(a[0]); System.out.println(a[3]);",
    );
    assert_eq!(out, vec!["1", "0"]);
}

#[test]
fn arrays_binary_search_found() {
    let out = run_main(
        "int[] a = {1, 3, 5}; System.out.println(java.util.Arrays.binarySearch(a, 3));",
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn arrays_binary_search_missing() {
    let out = run_main(
        "int[] a = {1, 3, 5}; System.out.println(java.util.Arrays.binarySearch(a, 4) < 0);",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arrays_hash_code_consistent() {
    let out = run_main(
        "int[] a = {1, 2}; System.out.println(java.util.Arrays.hashCode(a) == java.util.Arrays.hashCode(a));",
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn arrays_to_string_format() {
    let out = run_main(
        "String s = java.util.Arrays.toString(new int[]{1, 2}); System.out.println(s);",
    );
    assert_eq!(out, vec!["[1, 2]"]);
}

#[test]
fn arrays_parallel_sort_subrange() {
    let out = run_main(
        "int[] a = {9, 3, 1, 8}; java.util.Arrays.parallelSort(a, 1, 3); System.out.println(a[1]); System.out.println(a[2]);",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn arrays_compare_unsigned_bytes() {
    let out = run_main(
        "byte[] a = {(byte) 200}; byte[] b = {(byte) 100}; System.out.println(java.util.Arrays.compareUnsigned(a, b) > 0);",
    );
    assert_eq!(out, vec!["true"]);
}

