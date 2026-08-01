use crate::helpers::run_prints;

#[test]
fn test_int_array_basic_indexing_and_size() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3, 4)
            println(nums.size)
            println(nums[0])
            println(nums[3])
        }
    "#,
    );
    assert_eq!(out, &["4", "1", "4"]);
}

#[test]
fn test_int_array_factory_lambda() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = IntArray(4) { idx -> idx * idx }
            println(nums.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["0,1,4,9"]);
}

#[test]
fn test_int_array_sum_and_average() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3, 4)
            println(nums.sum())
            println(nums.average())
        }
    "#,
    );
    assert_eq!(out, &["10", "2.5"]);
}

#[test]
fn test_int_array_min_max() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(9, 1, 4, -2, 8)
            println(nums.minOrNull())
            println(nums.maxOrNull())
        }
    "#,
    );
    assert_eq!(out, &["-2", "9"]);
}

#[test]
fn test_int_array_copy_of_keeps_values() {
    let out = run_prints(
        r#"
        fun main() {
            val base = intArrayOf(1, 2, 3)
            val copy = base.copyOf()
            copy[0] = 10
            println(base[0])
            println(copy[0])
            println(base.size)
        }
    "#,
    );
    assert_eq!(out, &["1", "10", "3"]);
}

#[test]
fn test_int_array_copy_of_with_new_size() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3)
            val grown = nums.copyOf(5)
            val shrunk = nums.copyOf(2)
            println(grown.joinToString(","))
            println(shrunk.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,0,0", "1,2"]);
}

#[test]
fn test_int_array_copy_of_range_subset() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(10, 20, 30, 40, 50)
            val mid = nums.copyOfRange(1, 4)
            val empty = nums.copyOfRange(4, 4)
            println(mid.joinToString(","))
            println(empty.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["20,30,40", ""]);
}

#[test]
fn test_int_array_fill_whole_array() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = IntArray(4)
            nums.fill(7)
            println(nums.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["7,7,7,7"]);
}

#[test]
fn test_int_array_fill_with_range() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 1, 1, 1, 1)
            nums.fill(9, 1, 4)
            println(nums.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,9,9,9,1"]);
}

#[test]
fn test_int_array_sort_in_place() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(3, 1, 4, 1, 5, 9)
            nums.sort()
            println(nums.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,1,3,4,5,9"]);
}

#[test]
fn test_int_array_sort_descending_copy() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(3, 1, 4, 2)
            val desc = nums.sortedArrayDescending()
            println(desc.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["4,3,2,1"]);
}

#[test]
fn test_int_array_binary_search_found_and_missing() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 3, 5, 7)
            println(nums.binarySearch(5))
            println(nums.binarySearch(4))
        }
    "#,
    );
    assert_eq!(out, &["2", "-3"]);
}

#[test]
fn test_int_array_binary_search_range_start() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3, 4, 5, 6)
            println(nums.binarySearch(4, 1, 4))
            println(nums.binarySearch(4, 0, 3))
        }
    "#,
    );
    assert_eq!(out, &["3", "-5"]);
}

#[test]
fn test_int_array_get_or_else_get_or_null() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(9, 8, 7)
            println(nums.getOrElse(1) { -1 })
            println(nums.getOrNull(10) ?: -1)
        }
    "#,
    );
    assert_eq!(out, &["8", "-1"]);
}

#[test]
fn test_array_with_function_mapping() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3).map { it * 2 }.toIntArray()
            println(nums.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4,6"]);
}

#[test]
fn test_array_reverse_in_place_and_copy() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3, 4)
            nums.reverse()
            println(nums.joinToString(","))
            val back = nums.reversedArray()
            println(back.joinToString(","))
            println(nums[0] + back[0])
        }
    "#,
    );
    assert_eq!(out, &["4,3,2,1", "1,2,3,4", "5"]);
}

#[test]
fn test_int_array_to_typed_and_back() {
    let out = run_prints(
        r#"
        fun main() {
            val primitive = intArrayOf(1, 2, 3)
            val boxed = primitive.toTypedArray().toIntArray()
            println(boxed.joinToString(","))
            println(primitive.contentEquals(boxed))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "true"]);
}

#[test]
fn test_array_of_char_and_to_char_array_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            val text = "kot"
            val chars = text.toCharArray()
            println(chars.joinToString(","))
            println(String(chars))
        }
    "#,
    );
    assert_eq!(out, &["k,o,t", "kot"]);
}

#[test]
fn test_char_array_sorted_copy() {
    let out = run_prints(
        r#"
        fun main() {
            val chars = charArrayOf('c', 'a', 'b')
            val sorted = chars.sortedArray()
            println(sorted.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["a,b,c"]);
}

#[test]
fn test_byte_array_fill_and_sum() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = byteArrayOf(1, 2, 3)
            bytes.fill(9.toByte(), 0, 2)
            var total = 0
            for (b in bytes) {
                total += b.toInt()
            }
            println(bytes.joinToString(","))
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["9,9,3", "21"]);
}

#[test]
fn test_array_content_equality_for_object_arrays() {
    let out = run_prints(
        r#"
        fun main() {
            val a = arrayOf("x", "y")
            val b = arrayOf("x", "y")
            val c = arrayOf("x", "z")
            println(a.contentEquals(b))
            println(a.contentEquals(c))
            println(a.contentToString())
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "[x, y]"]);
}

#[test]
fn test_nested_array_deep_equality() {
    let out = run_prints(
        r#"
        fun main() {
            val a = arrayOf(arrayOf(1, 2), arrayOf(3))
            val b = arrayOf(arrayOf(1, 2), arrayOf(3))
            val c = arrayOf(arrayOf(1, 2), arrayOf(4))
            println(a.contentDeepEquals(b))
            println(a.contentDeepEquals(c))
            println(a.contentDeepToString())
        }
    "#,
    );
    assert_eq!(out, &["true", "false", "[[1, 2], [3]]"]);
}

#[test]
fn test_array_slice_array() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(10, 20, 30, 40, 50)
            val slice = nums.sliceArray(1..3)
            println(slice.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["20,30,40"]);
}

#[test]
fn test_array_filter_to_intarray() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3, 4, 5, 6)
            val evens = nums.filter { it % 2 == 0 }.toIntArray()
            println(evens.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,4,6"]);
}

#[test]
fn test_array_any_all_none_count() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3, 4)
            println(nums.any { it > 3 })
            println(nums.all { it > 0 })
            println(nums.none { it < 0 })
            println(nums.count { it % 2 == 0 })
        }
    "#,
    );
    assert_eq!(out, &["true", "true", "true", "2"]);
}

#[test]
fn test_array_find_first_last_or_null() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(2, 4, 6)
            println(nums.find { it > 3 } ?: -1)
            println(nums.findLast { it < 4 } ?: -1)
            println(nums.firstOrNull { it == 10 } ?: -1)
        }
    "#,
    );
    assert_eq!(out, &["4", "2", "-1"]);
}

#[test]
fn test_array_for_each_with_indexed_sink() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(4, 5, 6)
            var out = ""
            nums.forEachIndexed { index, value ->
                out += index.toString() + ":" + value.toString() + ";"
            }
            println(out)
        }
    "#,
    );
    assert_eq!(out, &["0:4;1:5;2:6;"]);
}

#[test]
fn test_array_distinct_and_distinct_by_key() {
    let out = run_prints(
        r#"
        fun main() {
            val items = arrayOf("aa", "ab", "b", "cc")
            val distinctByLen = items.distinctBy { it.length }
            println(distinctByLen.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["aa,b"]);
}

#[test]
fn test_array_out_of_bounds_raises() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3)
            try {
                println(nums[3])
            } catch (e: IndexOutOfBoundsException) {
                println("out_of_bounds")
            }
        }
    "#,
    );
    assert_eq!(out, &["out_of_bounds"]);
}

#[test]
fn test_array_set_and_retrieve_with_last_index() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = IntArray(3)
            nums[nums.lastIndex] = 99
            println(nums.last())
            println(nums.lastIndex)
        }
    "#,
    );
    assert_eq!(out, &["99", "2"]);
}

#[test]
fn test_array_is_empty_and_indices_contract() {
    let out = run_prints(
        r#"
        fun main() {
            val none = intArrayOf()
            println(none.isEmpty())
            println(none.isNotEmpty())
            println(none.indices.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["true", "false", ""]);
}

#[test]
fn test_array_first_or_null_on_empty() {
    let out = run_prints(
        r#"
        fun main() {
            val none = intArrayOf()
            println(none.firstOrNull() ?: "missing")
            println(none.lastOrNull() ?: "missing")
        }
    "#,
    );
    assert_eq!(out, &["missing", "missing"]);
}

#[test]
fn test_int_array_first_throws_on_empty() {
    let out = run_prints(
        r#"
        fun main() {
            val none = intArrayOf()
            try {
                println(none.first())
            } catch (e: NoSuchElementException) {
                println("missing")
            }
        }
    "#,
    );
    assert_eq!(out, &["missing"]);
}

#[test]
fn test_int_array_with_index_iteration() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(2, 4, 6)
            var trace = ""
            for ((idx, value) in nums.withIndex()) {
                trace += idx.toString() + ":" + value.toString() + ";"
            }
            println(trace)
        }
    "#,
    );
    assert_eq!(out, &["0:2;1:4;2:6;"]);
}

#[test]
fn test_array_as_list_roundtrip_is_copy() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3)
            val list = nums.toTypedArray().toList()
            val rebuilt = list.toIntArray()
            println(list.size)
            println(rebuilt.joinToString(","))
            println(list[1])
        }
    "#,
    );
    assert_eq!(out, &["3", "1,2,3", "2"]);
}

#[test]
fn test_concatenate_arrays_via_plus() {
    let out = run_prints(
        r#"
        fun main() {
            val a = intArrayOf(1, 2)
            val b = intArrayOf(3, 4)
            val c = a + b
            println(c.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3,4"]);
}

#[test]
fn test_array_fill_with_step() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = IntArray(6)
            for (i in nums.indices step 2) {
                nums[i] = i
            }
            println(nums.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["0,0,2,0,4,0"]);
}

#[test]
fn test_array_zip_with_prefix_and_length_mismatch() {
    let out = run_prints(
        r#"
        fun main() {
            val left = intArrayOf(1, 2, 3)
            val right = intArrayOf(10, 20)
            val pairs = left.zip(right.toTypedArray())
            val values = pairs.joinToString("|") { it.first.toString() + ":" + it.second.toString() }
            println(values)
        }
    "#,
    );
    assert_eq!(out, &["1:10|2:20"]);
}

#[test]
fn test_array_binary_search_with_comparator_like_transform() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 3, 5, 7, 9)
            println(nums.binarySearch(5))
            println(nums.binarySearch(6))
            println(nums.binarySearch(6, 0, 4))
        }
    "#,
    );
    assert_eq!(out, &["2", "-4", "-4"]);
}

#[test]
fn test_byte_array_join_and_sum() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = byteArrayOf(5, -1, 3)
            val shifted = bytes.map { it.toInt() + 1 }.toByteArray()
            var total = 0
            for (value in shifted) {
                total += value.toInt()
            }
            println(shifted.joinToString(","))
            println(total)
        }
    "#,
    );
    assert_eq!(out, &["6,0,4", "10"]);
}

#[test]
fn test_int_array_index_of_and_last_index_of() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(4, 1, 4, 2, 4)
            println(nums.indexOf(4))
            println(nums.lastIndexOf(4))
            println(nums.indexOf(9))
        }
    "#,
    );
    assert_eq!(out, &["0", "4", "-1"]);
}

#[test]
fn test_int_array_sort_range_only() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(9, 1, 8, 3, 2, 7)
            nums.sort(1, 5)
            println(nums.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["9,1,2,3,8,7"]);
}

#[test]
fn test_int_array_to_int_list_is_snapshot() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3)
            val list = nums.toTypedArray().toMutableList()
            list[0] = 9
            nums[1] = 5
            println(nums.joinToString(","))
            println(list.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,5,3", "9,2,3"]);
}

#[test]
fn test_array_content_hash_code_is_positive() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3)
            val nested = arrayOf(nums)
            println(nums.contentHashCode() > 0)
            println(nested.contentDeepHashCode() > 0)
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_byte_array_join_to_string_with_charset_preserves_signs() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = byteArrayOf(1, -2, 127, -128)
            println(bytes.joinToString("|"))
            val first = bytes[1].toInt()
            val last = bytes[3].toInt()
            println(first + last)
        }
    "#,
    );
    assert_eq!(out, &["1,-2,127,-128", "-130"]);
}

#[test]
fn test_array_of_nulls_defaults_to_null_and_can_mutate_indices() {
    let out = run_prints(
        r#"
        fun main() {
            val slots = arrayOfNulls<String>(3)
            val before = slots.joinToString(",") { it ?: "null" }
            slots[1] = "value"
            val after = slots.joinToString(",") { it ?: "null" }
            println(before)
            println(after)
            println(slots.count { it == null })
        }
    "#,
    );
    assert_eq!(out, &["null,null,null", "null,value,null", "2"]);
}

#[test]
fn test_int_array_fold_and_reduce_contracts() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3, 4)
            println(nums.reduce { acc, value -> acc + value })
            println(nums.fold(10) { acc, value -> acc - value })
            println(nums.fold("") { acc, value -> acc + value.toString() })
        }
    "#,
    );
    assert_eq!(out, &["10", "0", "1234"]);
}

#[test]
fn test_array_take_last_and_drop() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3, 4, 5)
            val tail = nums.takeLast(2)
            val mid = nums.drop(2)
            println(tail.joinToString(","))
            println(mid.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["4,5", "3,4,5"]);
}

#[test]
fn test_int_array_mutation_with_for_each_indexed() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3)
            nums.forEachIndexed { index, value ->
                nums[index] = value * 3
            }
            println(nums.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3,6,9"]);
}

#[test]
fn test_array_copy_of_range_out_of_bounds_throws() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = intArrayOf(1, 2, 3)
            try {
                nums.copyOfRange(-1, 2)
            } catch (e: IllegalArgumentException) {
                println("bad")
            }
        }
    "#,
    );
    assert_eq!(out, &["bad"]);
}
