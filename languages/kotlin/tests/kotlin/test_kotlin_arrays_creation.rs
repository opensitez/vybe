kotlin_run_test!(
    test_int_array_of_constructor_and_indexing,
    r#"
        fun main() {
            val values = intArrayOf(1, 2, 3, 4)
            println(values[0] + values[3])
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_byte_array_initializer_and_sum,
    r#"
        fun main() {
            val bytes = byteArrayOf(1, 2, 3)
            println(bytes[1] + bytes[2])
        }
    "#,
    &["5"]
);

kotlin_run_test!(
    test_boolean_array_fill_and_to_list_behavior,
    r#"
        fun main() {
            val bits = BooleanArray(4)
            bits[0] = true
            bits[2] = true
            println(bits.count { it })
        }
    "#,
    &["2"]
);

kotlin_run_test!(
    test_generic_array_of_nulls_with_runtime_type,
    r#"
        fun main() {
            val values = arrayOfNulls<String>(3)
            values[0] = "a"
            values[1] = "b"
            values[2] = "c"
            println(values.filterNotNull().joinToString(","))
        }
    "#,
    &["a,b,c"]
);

kotlin_run_test!(
    test_array_constructor_lambda_and_copy,
    r#"
        fun main() {
            val values = Array(4) { it * 2 }
            val copy = values.copyOf(2)
            println(copy.joinToString(","))
        }
    "#,
    &["0,2"]
);

kotlin_run_test!(
    test_primitive_to_list_and_join,
    r#"
        fun main() {
            val values = longArrayOf(4L, 8L, 12L)
            println(values.toList().joinToString(","))
        }
    "#,
    &["4,8,12"]
);

kotlin_run_test!(
    test_copy_of_range_with_offseted_start,
    r#"
        fun main() {
            val values = intArrayOf(0, 1, 2, 3, 4, 5)
            val slice = values.copyOfRange(1, 4)
            println(slice.joinToString(","))
        }
    "#,
    &["1,2,3"]
);

kotlin_run_test!(
    test_array_of_any_casting_and_mutation,
    r#"
        fun main() {
            val mix = arrayOf<Any>("x", 1, true)
            val tail = mix.drop(1)
            println(tail.joinToString("|"))
        }
    "#,
    &["1|true"]
);

kotlin_run_test!(
    test_char_array_to_string_roundtrip,
    r#"
        fun main() {
            val chars = charArrayOf('k', 'o', 't', 'l', 'i', 'n')
            val joined = chars.concatToString()
            println(joined)
            println(joined.toCharArray().size)
        }
    "#,
    &["kotlin", "6"]
);

kotlin_run_test!(
    test_double_array_range_with_fill,
    r#"
        fun main() {
            val values = DoubleArray(3)
            values.fill(2.5)
            println(values.sum())
        }
    "#,
    &["7.5"]
);

kotlin_run_test!(
    test_object_array_as_mutable_list_roundtrip,
    r#"
        fun main() {
            val values = arrayOf("a", "b", "c")
            val copy = values.toMutableList()
            copy.add("d")
            println(copy.joinToString(";"))
        }
    "#,
    &["a;b;c;d"]
);
