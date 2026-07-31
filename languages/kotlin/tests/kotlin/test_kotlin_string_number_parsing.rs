kotlin_run_cases! {
    test_parse_int_success_and_failure => (r##"
        fun main() {
            println("12".toInt())
            println("999".toIntOrNull())
            println("x12".toIntOrNull())
            println("x12".toIntOrNull()?.toString() ?: "null")
        }
    "##, &[
        "12",
        "999",
        "null",
        "null",
    ]),
    test_parse_long_with_sign => (r##"
        fun main() {
            println("-12".toLong())
            println("+12".toLong())
            println("12L".toLongOrNull())
        }
    "##, &[
        "-12",
        "12",
        "null",
    ]),
    test_parse_double_and_float => (r##"
        fun main() {
            println("12.75".toDouble())
            println("-0.5".toFloat())
            println("nan".toDoubleOrNull()?.isNaN() ?: false)
            println("bad".toDoubleOrNull())
        }
    "##, &[
        "12.75",
        "-0.5",
        "true",
        "null",
    ]),
    test_parse_boolean_values => (r##"
        fun main() {
            println("true".toBoolean())
            println("false".toBoolean())
            println("TRUE".toBoolean())
            println("true".toBooleanStrictOrNull())
            println("TRUE".toBooleanStrictOrNull())
            println("x".toBooleanStrictOrNull())
        }
    "##, &[
        "true",
        "false",
        "false",
        "true",
        "null",
        "null",
    ]),
    test_parse_radix_values => (r##"
        fun main() {
            println("ff".toInt(16))
            println("11".toInt(2))
            println("77".toInt(8))
            println("123".toIntOrNull(37))
        }
    "##, &[
        "255",
        "3",
        "63",
        "null",
    ]),
    test_parse_byte_bounds => (r##"
        fun main() {
            println("127".toByteOrNull())
            println("128".toByteOrNull())
            println("-128".toByte())
            println("-129".toByteOrNull())
        }
    "##, &[
        "127",
        "null",
        "-128",
        "null",
    ]),
    test_parse_boolean_or_default => (r##"
        fun main() {
            val v = "x".toBooleanOrNull()
            println(v ?: false)
            val n = "x".toIntOrNull()
            println(n ?: 0)
        }
    "##, &[
        "false",
        "0",
    ]),
    test_parse_nullable_coalesce => (r##"
        fun main() {
            val raw = listOf("10", "x", "20")
            val sum = raw.mapNotNull { it.toIntOrNull() }.sum()
            println(sum)
            println(raw.mapNotNull { it.toIntOrNull(16) }.size)
        }
    "##, &[
        "30",
        "0",
    ]),
}
