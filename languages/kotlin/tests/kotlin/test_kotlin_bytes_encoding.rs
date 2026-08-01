use crate::helpers::run_prints;

#[test]
fn test_byte_array_factory_size_and_values() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = byteArrayOf(7, 8, 9, 10)
            println(bytes.size)
            println(bytes.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["4", "7,8,9,10"]);
}

#[test]
fn test_string_to_byte_array_round_trip_ascii() {
    let out = run_prints(
        r#"
        fun main() {
            val source = "Kotlin"
            val bytes = source.toByteArray()
            val value = String(bytes)
            println(bytes.size)
            println(value)
        }
    "#,
    );
    assert_eq!(out, &["6", "Kotlin"]);
}

#[test]
fn test_string_to_byte_array_round_trip_utf8_unicode() {
    let out = run_prints(
        r#"
        fun main() {
            val source = "hé"
            val bytes = source.toByteArray(Charsets.UTF_8)
            val value = String(bytes, Charsets.UTF_8)
            println(bytes.size)
            println(value)
            println(bytes.first())
        }
    "#,
    );
    assert_eq!(out, &["3", "hé", "-61"]);
}

#[test]
fn test_string_to_byte_array_iso_8859_1() {
    let out = run_prints(
        r#"
        fun main() {
            val source = "Cafe"
            val bytes = source.toByteArray(Charsets.ISO_8859_1)
            val value = String(bytes, Charsets.ISO_8859_1)
            println(bytes.size)
            println(value)
            println(bytes.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["4", "Cafe", "67,97,102,101"]);
}

#[test]
fn test_string_to_byte_array_us_ascii() {
    let out = run_prints(
        r#"
        fun main() {
            val source = "abc123"
            val bytes = source.toByteArray(Charsets.US_ASCII)
            val value = String(bytes, Charsets.US_ASCII)
            println(bytes.size)
            println(value)
            println(bytes.sum())
        }
    "#,
    );
    assert_eq!(out, &["6", "abc123", "594"]);
}

#[test]
fn test_byte_array_fill_to_value() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = ByteArray(5)
            bytes.fill(7)
            println(bytes.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["7,7,7,7,7"]);
}

#[test]
fn test_byte_array_fill_with_range() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = ByteArray(8)
            bytes.fill(1, fromIndex = 2, toIndex = 6)
            println(bytes.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["0,0,1,1,1,1,0,0"]);
}

#[test]
fn test_byte_array_mutable_set_access() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = byteArrayOf(10, 20, 30)
            bytes[1] = 55
            println(bytes[0])
            println(bytes[1])
            println(bytes[2])
            println(bytes.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["10", "55", "30", "10,55,30"]);
}

#[test]
fn test_byte_array_copy_of_exact_size() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = byteArrayOf(1, 2, 3)
            val copied = bytes.copyOf()
            println(copied.joinToString(","))
            println(copied.contentEquals(bytes))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "true"]);
}

#[test]
fn test_byte_array_copy_of_expands_with_zeros() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = byteArrayOf(1, 2, 3)
            val expanded = bytes.copyOf(5)
            println(expanded.joinToString(","))
            println(expanded.size)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,0,0", "5"]);
}

#[test]
fn test_byte_array_copy_of_truncates() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = byteArrayOf(1, 2, 3, 4, 5)
            val truncated = bytes.copyOf(3)
            println(truncated.joinToString(","))
            println(truncated.size)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "3"]);
}

#[test]
fn test_byte_array_copy_of_range_middle_slice() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = byteArrayOf(10, 20, 30, 40, 50)
            val middle = bytes.copyOfRange(1, 4)
            println(middle.joinToString(","))
            println(middle.size)
        }
    "#,
    );
    assert_eq!(out, &["20,30,40", "3"]);
}

#[test]
fn test_byte_array_plus_operator() {
    let out = run_prints(
        r#"
        fun main() {
            val first = byteArrayOf(1, 2)
            val second = byteArrayOf(3, 4, 5)
            val both = first + second
            println(both.joinToString(","))
            println(first.size)
            println(second.size)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,4,5", "2", "3"]);
}

#[test]
fn test_byte_array_reversed_array() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(1, 3, 5)
            val reversed = values.reversedArray()
            println(reversed.joinToString(","))
            println(values.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["5,3,1", "1,3,5"]);
}

#[test]
fn test_byte_array_sum_and_reduce() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(1, 2, 3)
            println(values.sum())
            println(values.reduce { acc, value -> acc + value })
            println(values.fold(0) { acc, value -> acc + value })
        }
    "#,
    );
    assert_eq!(out, &["6", "6", "6"]);
}

#[test]
fn test_byte_array_content_equals_and_not_equals() {
    let out = run_prints(
        r#"
        fun main() {
            val a = byteArrayOf(1, 2, 3)
            val b = byteArrayOf(1, 2, 3)
            val c = byteArrayOf(3, 2, 1)
            println(a.contentEquals(b))
            println(a.contentEquals(c))
            println(a.contentEquals(a))
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "true"]);
}

#[test]
fn test_byte_array_map_transformation() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(1, 2, 3)
            val doubled = values.map { it * 2 }
            println(doubled.joinToString(","))
            println(doubled.sum())
        }
    "#,
    );
    assert_eq!(out, &["2,4,6", "12"]);
}

#[test]
fn test_byte_array_filter_and_all_any() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(1, 2, 3, 4, 5)
            println(values.filter { it % 2 == 0 }.joinToString(","))
            println(values.any { it > 4 })
            println(values.all { it > 0 })
            println(values.none { it < 0 })
        }
    "#,
    );
    assert_eq!(out, &["2,4", "true", "true", "true"]);
}

#[test]
fn test_byte_array_count_and_first_last() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(8, 0, 8, 4)
            println(values.count { it == 8 })
            println(values.first())
            println(values.last())
        }
    "#,
    );
    assert_eq!(out, &["2", "8", "4"]);
}

#[test]
fn test_byte_array_take_and_drop() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(1, 2, 3, 4)
            println(values.take(3).joinToString(","))
            println(values.takeLast(2).joinToString(","))
            println(values.drop(1).joinToString(","))
            println(values.dropLast(1).joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "3,4", "2,3,4", "1,2,3"]);
}

#[test]
fn test_byte_array_zip_addition() {
    let out = run_prints(
        r#"
        fun main() {
            val left = byteArrayOf(1, 2, 3)
            val right = byteArrayOf(4, 5, 6)
            val zipped = left.zip(right) { a, b -> a + b }
            println(zipped.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["5,7,9"]);
}

#[test]
fn test_byte_array_join_to_string_formatting() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(3, 1, 4)
            println(values.joinToString("|"))
            println(values.joinToString(",", prefix = "[", postfix = "]"))
            println(values.joinToString(";", transform = { it.toString() }))
        }
    "#,
    );
    assert_eq!(out, &["3|1|4", "[3,1,4]", "3;1;4"]);
}

#[test]
fn test_byte_array_iterate_with_indices() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(4, 5, 6)
            var out = ""
            for (item in values.withIndex()) {
                out += item.index.toString() + ":" + item.value.toString() + ";"
            }
            println(out)
        }
    "#,
    );
    assert_eq!(out, &["0:4;1:5;2:6;"]);
}

#[test]
fn test_byte_array_loop_sum_accumulator() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(2, 4, 6)
            var total = 0
            for (value in values) {
                total += value.toInt()
            }
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_byte_array_as_sequence_map_to_sum() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(2, 3, 4)
            val total = values.asSequence().map { it.toInt() }.sum()
            println(total)
            println(values.asSequence().map { it + 1 }.toList().joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["9", "3,4,5"]);
}

#[test]
fn test_string_from_ascii_bytes_empty() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = byteArrayOf()
            val value = String(bytes)
            println(bytes.size)
            println(value.isEmpty())
        }
    "#,
    );
    assert_eq!(out, &["0", "true"]);
}

#[test]
fn test_string_from_byte_array_with_nulls() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = byteArrayOf(78, 117, 108, 108, 0, 65)
            val value = String(bytes)
            println(value.length)
            println(value[0])
            println(value[4].code)
        }
    "#,
    );
    assert_eq!(out, &["6", "N", "0"]);
}

#[test]
fn test_byte_array_empty_copy_of_range_behavior() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = byteArrayOf(1, 2, 3)
            val empty = bytes.copyOfRange(1, 1)
            println(empty.size)
            println(empty.isEmpty())
        }
    "#,
    );
    assert_eq!(out, &["0", "true"]);
}

#[test]
fn test_byte_array_to_hex_strings() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(0, 10, 15, 127, -1)
            val hex = values.joinToString(",") {
                val u = it.toInt() and 0xFF
                u.toString(16).padStart(2, '0')
            }
            println(hex)
        }
    "#,
    );
    assert_eq!(out, &["00,0a,0f,7f,ff"]);
}

#[test]
fn test_byte_array_find_first_greater_than_threshold() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(4, 6, 8, 9)
            var found = -1
            for (value in values) {
                if (value > 7) {
                    found = value.toInt()
                    break
                }
            }
            println(found)
        }
    "#,
    );
    assert_eq!(out, &["8"]);
}

#[test]
fn test_byte_array_count_even_values() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(1, 2, 3, 4, 5, 6)
            println(values.count { it % 2 == 0 })
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_byte_array_plus_with_empty_array() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(1, 2, 3) + byteArrayOf()
            println(values.joinToString(","))
            val empty = byteArrayOf() + byteArrayOf(9, 10)
            println(empty.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "9,10"]);
}

#[test]
fn test_byte_array_copy_of_range_out_of_bounds_throws() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(1, 2, 3)
            try {
                values.copyOfRange(4, 6)
                println("ok")
            } catch (e: Exception) {
                println(e::class.simpleName)
            }
        }
    "#,
    );
    assert_eq!(out, &["IndexOutOfBoundsException"]);
}

#[test]
fn test_byte_array_set_out_of_bounds_throws() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(1, 2, 3)
            try {
                values[7] = 9
                println("ok")
            } catch (e: Exception) {
                println(e::class.simpleName)
            }
        }
    "#,
    );
    assert_eq!(out, &["ArrayIndexOutOfBoundsException"]);
}

#[test]
fn test_byte_array_to_char_list_and_string() {
    let out = run_prints(
        r#"
        fun main() {
            val values = "AZ09".toByteArray()
            val chars = values.map { it.toInt().toChar() }.joinToString(",")
            println(chars)
            val rebuilt = String(byteArrayOf(65, 90, 48, 57))
            println(rebuilt)
        }
    "#,
    );
    assert_eq!(out, &["A,Z,0,9", "AZ09"]);
}

#[test]
fn test_byte_array_sum_in_long_accumulator() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(100, 101, 102)
            var total: Long = 0
            for (value in values) {
                total += value.toLong()
            }
            println(total)
            println(total > 300)
        }
    "#,
    );
    assert_eq!(out, &["303", "true"]);
}

#[test]
fn test_byte_array_to_typed_array_to_mutable_list() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(1, 2, 3).toTypedArray().toMutableList()
            values.add(4)
            println(values.joinToString(","))
            println(values.size)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,4", "4"]);
}

#[test]
fn test_byte_array_filter_not_zero() {
    let out = run_prints(
        r#"
        fun main() {
            val values = byteArrayOf(0, 1, 0, 2, 3)
            val filtered = values.filterNot { it == 0 }
            println(filtered.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3"]);
}

#[test]
fn test_byte_array_round_trip_emoji_utf8() {
    let out = run_prints(
        r#"
        fun main() {
            val source = "🙂"
            val bytes = source.toByteArray(Charsets.UTF_8)
            val value = String(bytes, Charsets.UTF_8)
            println(bytes.size)
            println(value)
            println(value == source)
        }
    "#,
    );
    assert_eq!(out, &["4", "🙂", "true"]);
}
