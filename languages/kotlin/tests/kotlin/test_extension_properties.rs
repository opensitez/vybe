use crate::helpers::run_prints;

#[test]
fn test_extension_property_int_is_even() {
    let out = run_prints(
        r#"
        val Int.isEven: Boolean get() = this % 2 == 0
        fun main() {
            println(4.isEven)
            println(5.isEven)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_extension_property_int_double() {
    let out = run_prints(
        r#"
        val Int.doubled: Int get() = this * 2
        fun main() {
            println(7.doubled)
        }
    "#,
    );
    assert_eq!(out, &["14"]);
}

#[test]
fn test_extension_property_int_triple() {
    let out = run_prints(
        r#"
        val Int.tripled: Int get() = this * 3
        fun main() {
            println(4.tripled)
        }
    "#,
    );
    assert_eq!(out, &["12"]);
}

#[test]
fn test_extension_property_string_first_char() {
    let out = run_prints(
        r#"
        val String.firstChar: Char get() = this[0]
        fun main() {
            println("hello".firstChar)
        }
    "#,
    );
    assert_eq!(out, &["h"]);
}

#[test]
fn test_extension_property_string_last_char() {
    let out = run_prints(
        r#"
        val String.lastChar: Char get() = this[this.length - 1]
        fun main() {
            println("rust".lastChar)
        }
    "#,
    );
    assert_eq!(out, &["t"]);
}

#[test]
fn test_extension_property_string_reversed_label() {
    let out = run_prints(
        r#"
        val String.reversedLabel: String get() = this.reversed()
        fun main() {
            println("abc".reversedLabel)
        }
    "#,
    );
    assert_eq!(out, &["cba"]);
}

#[test]
fn test_extension_property_string_length_without_spaces() {
    let out = run_prints(
        r#"
        val String.trimmedLength: Int get() = this.trim().length
        fun main() {
            println("  a b  ".trimmedLength)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_extension_property_list_is_empty() {
    let out = run_prints(
        r#"
        val <T> List<T>.isNoItems: Boolean get() = this.isEmpty()
        fun main() {
            println(listOf<Int>().isNoItems)
            println(listOf(1, 2).isNoItems)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_extension_property_list_size_text() {
    let out = run_prints(
        r#"
        val List<Int>.sizeText: String get() = when (size) {
            0 -> "empty"
            1 -> "single"
            else -> "many"
        }
        fun main() {
            println(listOf<Int>().sizeText)
            println(listOf(1).sizeText)
            println(listOf(1, 2).sizeText)
        }
    "#,
    );
    assert_eq!(out, &["empty", "single", "many"]);
}

#[test]
fn test_extension_property_nullable_string_or_empty() {
    let out = run_prints(
        r#"
        val String?.valueOrDash: String get() = this ?: "-"
        fun main() {
            val a: String? = null
            val b: String? = "ok"
            println(a.valueOrDash)
            println(b.valueOrDash)
        }
    "#,
    );
    assert_eq!(out, &["-", "ok"]);
}

#[test]
fn test_extension_property_nullable_string_length() {
    let out = run_prints(
        r#"
        val String?.orZeroLength: Int get() = this?.length ?: 0
        fun main() {
            val x: String? = null
            val y: String? = "kotlin"
            println(x.orZeroLength)
            println(y.orZeroLength)
        }
    "#,
    );
    assert_eq!(out, &["0", "6"]);
}

#[test]
fn test_extension_property_map_is_present() {
    let out = run_prints(
        r#"
        val <K, V> Map<K, V>.isPresent: Boolean get() = !isEmpty()
        fun main() {
            println(mapOf<String, Int>().isPresent)
            println(mapOf("a" to 1).isPresent)
        }
    "#,
    );
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_extension_property_pair_left() {
    let out = run_prints(
        r#"
        val <A, B> Pair<A, B>.left: A get() = first
        fun main() {
            println(Pair("a", 1).left)
        }
    "#,
    );
    assert_eq!(out, &["a"]);
}

#[test]
fn test_extension_property_pair_right() {
    let out = run_prints(
        r#"
        val <A, B> Pair<A, B>.rightText: String get() = second.toString()
        fun main() {
            println(Pair("a", 2).rightText)
        }
    "#,
    );
    assert_eq!(out, &["2"]);
}

#[test]
fn test_extension_property_boolean_as_int() {
    let out = run_prints(
        r#"
        val Boolean.asInt: Int get() = if (this) 1 else 0
        fun main() {
            println(true.asInt)
            println(false.asInt)
        }
    "#,
    );
    assert_eq!(out, &["1", "0"]);
}

#[test]
fn test_extension_property_person_full_name() {
    let out = run_prints(
        r#"
        class Person(val first: String, val last: String)
        val Person.fullName: String get() = "$first $last"
        fun main() {
            println(Person("Ada", "Lovelace").fullName)
        }
    "#,
    );
    assert_eq!(out, &["Ada Lovelace"]);
}

#[test]
fn test_extension_property_person_age_group() {
    let out = run_prints(
        r#"
        class Person(val age: Int)
        val Person.group: String get() = if (age < 18) "minor" else "adult"
        fun main() {
            println(Person(12).group)
            println(Person(28).group)
        }
    "#,
    );
    assert_eq!(out, &["minor", "adult"]);
}

#[test]
fn test_extension_property_range_is_small() {
    let out = run_prints(
        r#"
        val IntRange.isSmall: Boolean get() = this.count() <= 3
        fun main() {
            println((1..3).isSmall)
            println((1..5).isSmall)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_extension_property_char_uppercase_flag() {
    let out = run_prints(
        r#"
        val Char.isUpper: Boolean get() = this in 'A'..'Z'
        fun main() {
            println('R'.isUpper)
            println('t'.isUpper)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_extension_property_char_digit_flag() {
    let out = run_prints(
        r#"
        val Char.isDigitAscii: Boolean get() = this in '0'..'9'
        fun main() {
            println('3'.isDigitAscii)
            println('x'.isDigitAscii)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_extension_property_double_rounded() {
    let out = run_prints(
        r#"
        val Double.roundToEven: Int get() = kotlin.math.roundToInt(this)
        fun main() {
            println(3.4.roundToEven)
            println(2.6.roundToEven)
        }
    "#,
    );
    assert_eq!(out, &["3", "3"]);
}

#[test]
fn test_extension_property_set_of_int_min() {
    let out = run_prints(
        r#"
        val Set<Int>.minOrMinusOne: Int get() = this.minOrNull() ?: -1
        fun main() {
            println(setOf<Int>().minOrMinusOne)
            println(setOf(9, 2, 5).minOrMinusOne)
        }
    "#,
    );
    assert_eq!(out, &["-1", "2"]);
}

#[test]
fn test_extension_property_array_first_or_empty() {
    let out = run_prints(
        r#"
        val Array<Int>.firstOrNullSafe: Int get() = this.firstOrNull() ?: -1
        fun main() {
            println(intArrayOf().firstOrNullSafe)
            println(intArrayOf(1, 4).firstOrNullSafe)
        }
    "#,
    );
    assert_eq!(out, &["-1", "1"]);
}

#[test]
fn test_extension_property_boolean_to_word() {
    let out = run_prints(
        r#"
        val Boolean.word: String get() = if (this) "yes" else "no"
        fun main() {
            println(true.word)
            println(false.word)
        }
    "#,
    );
    assert_eq!(out, &["yes", "no"]);
}

#[test]
fn test_extension_property_array_as_text() {
    let out = run_prints(
        r#"
        val IntArray.totalText: String get() = this.joinToString(",")
        fun main() {
            println(intArrayOf(1, 2, 3).totalText)
        }
    "#,
    );
    assert_eq!(out, &["1,2,3"]);
}

#[test]
fn test_extension_property_nested_list_count() {
    let out = run_prints(
        r#"
        val List<List<Int>>.flattenedCount: Int get() = this.sumOf { it.size }
        fun main() {
            println(listOf(listOf(1, 2), listOf(3)).flattenedCount)
        }
    "#,
    );
    assert_eq!(out, &["3"]);
}

#[test]
fn test_extension_property_byte_is_zero() {
    let out = run_prints(
        r#"
        val Byte.isZero: Boolean get() = this == 0.toByte()
        fun main() {
            println(0.toByte().isZero)
            println(2.toByte().isZero)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_extension_property_string_is_numeric() {
    let out = run_prints(
        r#"
        val String.isNumeric: Boolean get() = this.all { ch -> ch in '0'..'9' }
        fun main() {
            println("123".isNumeric)
            println("12a3".isNumeric)
        }
    "#,
    );
    assert_eq!(out, &["true", "false"]);
}

#[test]
fn test_extension_property_range_span() {
    let out = run_prints(
        r#"
        val IntRange.span: Int get() = this.last - this.first
        fun main() {
            println((1..5).span)
        }
    "#,
    );
    assert_eq!(out, &["4"]);
}

#[test]
fn test_extension_property_list_non_empty() {
    let out = run_prints(
        r#"
        val List<*>.isNonEmpty: Boolean get() = this.isNotEmpty()
        fun main() {
            println(listOf<Int>().isNonEmpty)
            println(listOf(1, 2).isNonEmpty)
        }
    "#,
    );
    assert_eq!(out, &["false", "true"]);
}
