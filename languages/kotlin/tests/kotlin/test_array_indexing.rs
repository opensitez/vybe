kotlin_run_test!(
    test_int_array_get,
    r#"fun main() { val a = intArrayOf(1, 2, 3); println(a[1]) }"#,
    &["2"]
);

kotlin_run_test!(
    test_int_array_set,
    r#"fun main() { val a = intArrayOf(1, 2, 3); a[1] = 9; println(a.joinToString(",")) }"#,
    &["1,9,3"]
);

kotlin_run_test!(
    test_array_generic_get,
    r#"fun main() { val a = arrayOf("a", "b", "c"); println(a[2]) }"#,
    &["c"]
);

kotlin_run_test!(
    test_array_set_generic,
    r#"fun main() { val a = arrayOf("x", "y"); a[0] = "z"; println(a[0]) }"#,
    &["z"]
);

kotlin_run_test!(
    test_char_array_get,
    r#"fun main() { val a = charArrayOf('a', 'b', 'c'); println(a[0]) }"#,
    &["a"]
);

kotlin_run_test!(
    test_boolean_array_set,
    r#"fun main() { val a = booleanArrayOf(false, true); a[0] = true; println(a[0]) }"#,
    &["true"]
);

kotlin_run_test!(
    test_float_array_get,
    r#"fun main() { val a = floatArrayOf(1.5f, 2.5f); println(a[1]) }"#,
    &["2.5"]
);

kotlin_run_test!(
    test_list_get,
    r#"fun main() { val a = listOf(9, 8, 7); println(a[2]) }"#,
    &["7"]
);

kotlin_run_test!(
    test_mutable_list_set,
    r#"fun main() { val a = mutableListOf(1, 2, 3); a[1] = 5; println(a[1]) }"#,
    &["5"]
);

kotlin_run_test!(
    test_last_index,
    r#"fun main() { val a = intArrayOf(4, 5, 6); println(a[a.lastIndex]) }"#,
    &["6"]
);

kotlin_run_test!(
    test_array_first_last,
    r#"fun main() { val a = arrayOf("a", "b"); println(a.first() + a.last()) }"#,
    &["ab"]
);

kotlin_run_test!(
    test_array_subscript_after_math,
    r#"fun pick(values: IntArray, index: Int): Int = values[index]
fun main() { val a = intArrayOf(10, 20, 30); println(pick(a, 0 + 2)) }"#,
    &["30"]
);

kotlin_run_test!(
    test_multidim_array_get,
    r#"fun main() { val m = arrayOf(intArrayOf(1,2), intArrayOf(3,4)); println(m[1][0]) }"#,
    &["3"]
);

kotlin_run_test!(
    test_multidim_array_set,
    r#"fun main() { val m = arrayOf(intArrayOf(1,2), intArrayOf(3,4)); m[0][1] = 9; println(m[0][1]) }"#,
    &["9"]
);

kotlin_run_test!(
    test_string_char_at,
    r#"fun main() { val s = "kotlin"; println(s[1]) }"#,
    &["o"]
);

kotlin_run_test!(
    test_string_set_not_allowed,
    r#"fun main() {
        val s = "kotlin"
        try {
            // not actually executable by design
            val value = s[0]
            println(value)
        } catch (e: Exception) {
            println("ok")
        }
    }"#,
    &["k"]
);

kotlin_run_test!(
    test_safe_get_out_of_bounds,
    r#"fun main() { val a = intArrayOf(1, 2); try { println(a[5]) } catch (e: Exception) { println("err") } }"#,
    &["err"]
);

kotlin_run_test!(
    test_negative_index_out_of_bounds,
    r#"fun main() { val a = intArrayOf(1, 2); try { println(a[-1]) } catch (e: Exception) { println("err") } }"#,
    &["err"]
);

kotlin_run_test!(
    test_index_before_set,
    r#"fun main() { val a = intArrayOf(1, 2); a[0] = a[1] + 4; println(a.joinToString(",")) }"#,
    &["6,2"]
);

kotlin_run_test!(
    test_index_retain_order,
    r#"fun main() { val a = mutableListOf(3, 2, 1); val b = a[0] + a[2]; println(b) }"#,
    &["4"]
);

kotlin_run_test!(
    test_copy_from_range,
    r#"fun main() { val a = intArrayOf(1, 2, 3, 4); val b = a.copyOfRange(1, 3); println(b.joinToString(",")) }"#,
    &["2,3"]
);

kotlin_run_test!(
    test_fill_array_then_get,
    r#"fun main() { val a = IntArray(3); a.fill(5); println(a[2]) }"#,
    &["5"]
);

kotlin_run_test!(
    test_fill_element_at,
    r#"fun main() { val a = intArrayOf(1, 2, 3); java.util.Arrays.fill(a, 1, 2, 9); println(a[1]) }"#,
    &["9"]
);

kotlin_run_test!(
    test_indexed_get_on_range,
    r#"fun main() { val r = intArrayOf(5,6,7); println(r[1]) }"#,
    &["6"]
);

kotlin_run_test!(
    test_array_size_property,
    r#"fun main() { val a = IntArray(4); println(a.size) }"#,
    &["4"]
);

kotlin_run_test!(
    test_index_set_chained,
    r#"fun main() { val a = mutableListOf(1, 2, 3); a[a.lastIndex] = 9; println(a.last()) }"#,
    &["9"]
);

kotlin_run_test!(
    test_char_array_slice_to_string,
    r#"fun main() { val a = charArrayOf('x','y','z'); println(a.joinToString("")) }"#,
    &["xyz"]
);

kotlin_run_test!(
    test_string_index_last,
    r#"fun main() { val s = "hello"; println(s[s.lastIndex]) }"#,
    &["o"]
);

kotlin_run_test!(
    test_nested_list_mutation_by_index,
    r#"fun main() { val m = mutableListOf(mutableListOf(1,2), mutableListOf(3,4)); m[1][0] = 9; println(m[1][0]) }"#,
    &["9"]
);

kotlin_run_test!(
    test_assign_to_same_index,
    r#"fun main() { val a = IntArray(1); a[0] = a[0] + 1; println(a[0]) }"#,
    &["1"]
);

kotlin_run_test!(
    test_copy_of_then_index,
    r#"fun main() { val a = intArrayOf(9,8,7); val b = a.copyOf(5); b[4] = 1; println(b.size + b[4]) }"#,
    &["6"]
);

kotlin_run_test!(
    test_array_sort_and_index,
    r#"fun main() { val a = intArrayOf(4,1,3,2); java.util.Arrays.sort(a); println(a[0] + a[3]) }"#,
    &["5"]
);

kotlin_run_test!(
    test_map_key_lookup_by_index,
    r#"fun main() { val a = mapOf(1 to "a", 2 to "b"); println(a[2]) }"#,
    &["b"]
);

kotlin_run_test!(
    test_byte_array_indexing,
    r#"fun main() { val a = byteArrayOf(1,2,3); println(a[1]) }"#,
    &["2"]
);

kotlin_run_test!(
    test_long_array_reduce_index,
    r#"fun main() { val a = longArrayOf(2L, 4L, 6L); println(a[1] + a[2]) }"#,
    &["10"]
);

kotlin_run_test!(
    test_array_for_each_with_index,
    r#"fun main() {
        val a = intArrayOf(1, 2, 3)
        var sum = 0
        for (i in a.indices) { sum += a[i] }
        println(sum)
    }"#,
    &["6"]
);

kotlin_run_test!(
    test_array_with_for_each_indexed,
    r#"fun main() {
        val a = intArrayOf(1,2,3)
        var out = 0
        a.forEachIndexed { index, value -> if (index == 1) out = value }
        println(out)
    }"#,
    &["2"]
);

kotlin_run_test!(
    test_array_contains_indexed_access,
    r#"fun main() { val a = arrayOf("x", "y", "z"); println(a.indices.contains(1)) }"#,
    &["true"]
);

kotlin_run_test!(
    test_set_and_get_characters,
    r#"fun main() { val s = StringBuilder("abc"); s[1] = 'z'; println(s.toString()) }"#,
    &["azc"]
);
