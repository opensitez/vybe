use crate::helpers::run_prints;

#[test]
fn test_int_array_fill_and_sum() {
    let out = run_prints(r#"
        fun main() {
            val values = IntArray(3) { it + 1 }
            println(values.joinToString(","))
            println(values.sum())
        }
    "#);
    assert_eq!(out, &["1,2,3", "6"]);
}

#[test]
fn test_int_array_index_assignment() {
    let out = run_prints(r#"
        fun main() {
            val values = IntArray(3)
            values[1] = 9
            println(values[0])
            println(values[1])
            println(values[2])
        }
    "#);
    assert_eq!(out, &["0", "9", "0"]);
}

#[test]
fn test_int_array_copy_of() {
    let out = run_prints(r#"
        fun main() {
            val src = intArrayOf(1, 2, 3)
            val dst = src.copyOf(5)
            println(dst.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,3,0,0"]);
}

#[test]
fn test_int_array_copy_of_range() {
    let out = run_prints(r#"
        fun main() {
            val src = intArrayOf(1, 2, 3, 4)
            val dst = src.copyOfRange(1, 3)
            println(dst.joinToString(","))
        }
    "#);
    assert_eq!(out, &["2,3"]);
}

#[test]
fn test_int_array_get_or_else() {
    let out = run_prints(r#"
        fun main() {
            val src = intArrayOf(1, 2)
            println(src.getOrElse(0) { -1 })
            println(src.getOrElse(4) { -1 })
        }
    "#);
    assert_eq!(out, &["1", "-1"]);
}

#[test]
fn test_char_array_to_string() {
    let out = run_prints(r#"
        fun main() {
            val values = charArrayOf('k', 'o', 't')
            println(values.concatToString())
            println(values.size)
        }
    "#);
    assert_eq!(out, &["kot", "3"]);
}

#[test]
fn test_char_array_to_list() {
    let out = run_prints(r#"
        fun main() {
            val values = charArrayOf('a', 'b', 'c')
            println(values.toList().joinToString(","))
        }
    "#);
    assert_eq!(out, &["a,b,c"]);
}

#[test]
fn test_byte_array_sum_and_average() {
    let out = run_prints(r#"
        fun main() {
            val values = byteArrayOf(1, 2, 3)
            val sum = values.sum()
            val avg = values.average()
            println(sum)
            println(avg)
        }
    "#);
    assert_eq!(out, &["6", "2.0"]);
}

#[test]
fn test_boolean_array_any_all_none() {
    let out = run_prints(r#"
        fun main() {
            val values = booleanArrayOf(true, false, true)
            println(values.any())
            println(values.all { it })
            println(values.none { it == null })
        }
    "#);
    assert_eq!(out, &["true", "false", "true"]);
}

#[test]
fn test_double_array_join() {
    let out = run_prints(r#"
        fun main() {
            val values = doubleArrayOf(1.5, 2.5)
            println(values.joinToString(","))
            println(values.sum())
            println(values.average())
        }
    "#);
    assert_eq!(out, &["1.5,2.5", "4.0", "2.0"]);
}

#[test]
fn test_long_array_min_max() {
    let out = run_prints(r#"
        fun main() {
            val values = longArrayOf(10L, 3L, 7L)
            println(values.minOrNull())
            println(values.maxOrNull())
        }
    "#);
    assert_eq!(out, &["3", "10"]);
}

#[test]
fn test_float_array_conversion_to_ints() {
    let out = run_prints(r#"
        fun main() {
            val values = floatArrayOf(1.2f, 2.8f)
            val ints = values.map { it.toInt() }
            println(ints.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2"]);
}

#[test]
fn test_short_array_to_typed_list() {
    let out = run_prints(r#"
        fun main() {
            val values = shortArrayOf(4, 5)
            println(values.toTypedArray().joinToString(","))
            println(values.size)
        }
    "#);
    assert_eq!(out, &["4,5", "2"]);
}

#[test]
fn test_array_reversed_in_place() {
    let out = run_prints(r#"
        fun main() {
            val values = IntArray(4) { it + 1 }
            values.reverse()
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["4,3,2,1"]);
}

#[test]
fn test_array_sorting() {
    let out = run_prints(r#"
        fun main() {
            val values = intArrayOf(4, 1, 3, 2)
            values.sort()
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,3,4"]);
}

#[test]
fn test_array_slice_copy() {
    let out = run_prints(r#"
        fun main() {
            val values = intArrayOf(10, 20, 30, 40, 50)
            println(values.sliceArray(1..3).joinToString(","))
            println(values.slice(1..3).joinToString(","))
        }
    "#);
    assert_eq!(out, &["20,30,40", "20,30,40"]);
}

#[test]
fn test_array_content_equals() {
    let out = run_prints(r#"
        fun main() {
            val a = intArrayOf(1, 2)
            val b = intArrayOf(1, 2)
            println(a.contentEquals(b))
            println(a.contentHashCode() == b.contentHashCode())
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_array_set_all_via_fill() {
    let out = run_prints(r#"
        fun main() {
            val values = IntArray(3)
            values.fill(7)
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["7,7,7"]);
}

#[test]
fn test_u_int_array_size_and_access() {
    let out = run_prints(r#"
        fun main() {
            val values = uintArrayOf(1u, 2u, 3u)
            println(values.size)
            println(values[1])
        }
    "#);
    assert_eq!(out, &["3", "2"]);
}

#[test]
fn test_ubyte_array_to_ints() {
    let out = run_prints(r#"
        fun main() {
            val values = ubyteArrayOf(250u, 1u)
            val first = values[0].toInt()
            println(first)
            println(values.joinToString(","))
        }
    "#);
    assert_eq!(out, &["250", "250,1"]);
}

#[test]
fn test_ullong_array_to_string() {
    let out = run_prints(r#"
        fun main() {
            val values = ulongArrayOf(1UL, 2UL)
            println(values.joinToString(","))
            println(values[1])
        }
    "#);
    assert_eq!(out, &["1,2", "2"]);
}

#[test]
fn test_primitive_array_of_ranges() {
    let out = run_prints(r#"
        fun main() {
            val values = IntArray(5) { it * it }
            println(values.first())
            println(values.last())
            println(values[2])
        }
    "#);
    assert_eq!(out, &["0", "16", "4"]);
}

#[test]
fn test_boolean_array_to_int_mapping() {
    let out = run_prints(r#"
        fun main() {
            val values = booleanArrayOf(true, false, true)
            val mapped = values.map { if (it) 1 else 0 }
            println(mapped.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,0,1"]);
}

#[test]
fn test_char_array_indices_iteration() {
    let out = run_prints(r#"
        fun main() {
            val values = charArrayOf('a', 'b', 'c')
            var outText = ""
            for (i in values.indices) {
                outText += i.toString() + values[i]
            }
            println(outText)
        }
    "#);
    assert_eq!(out, &["0a1b2c"]);
}

#[test]
fn test_double_array_is_not_empty() {
    let out = run_prints(r#"
        fun main() {
            val values = doubleArrayOf()
            println(values.isEmpty())
            println(values.isNotEmpty())
        }
    "#);
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_array_iterate_sum_even_odd() {
    let out = run_prints(r#"
        fun main() {
            val values = longArrayOf(1, 2, 3, 4)
            var evens = 0
            var odds = 0
            for (v in values) {
                if (v % 2L == 0L) evens += 1 else odds += 1
            }
            println(evens)
            println(odds)
        }
    "#);
    assert_eq!(out, &["2", "2"]);
}

#[test]
fn test_array_transform_to_set() {
    let out = run_prints(r#"
        fun main() {
            val values = intArrayOf(1, 1, 2, 2, 3)
            val set = values.toMutableSet()
            println(set.joinToString(","))
            println(set.size)
        }
    "#);
    assert_eq!(out, &["1,2,3", "3"]);
}
