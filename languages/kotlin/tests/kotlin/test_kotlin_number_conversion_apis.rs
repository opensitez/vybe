kotlin_run_cases! {
    test_number_to_byte_and_short => (r##"
        fun main() {
            println(1234.7.toInt())
            println(129.toByte())
            println(128.toByte())
            println(32000.toShort())
            println(40000.toShort())
        }
    "##, &[
        "1234",
        "-126",
        "-128",
        "32000",
        "-25536",
    ]),
    test_number_to_int_from_float => (r##"
        fun main() {
            println(3.9.toInt())
            println((-3.9).toInt())
            println(3.4f.toInt())
            println((-3.4f).toInt())
        }
    "##, &[
        "3",
        "-3",
        "3",
        "-3",
    ]),
    test_number_to_long_and_float => (r##"
        fun main() {
            println(3_000_000_000L.toInt())
            println(10L.toDouble())
            println(42L.toFloat())
            println(42L.toByte())
        }
    "##, &[
        "-1294967296",
        "10000000000.0",
        "42.0",
        "42",
    ]),
    test_unsigned_roundtrip => (r##"
        fun main() {
            val u = (-1).toUInt()
            println(u)
            println(u.toInt())
            println(u.toLong())
            val w = u.toULong()
            println(w)
            println(w.toUInt())
        }
    "##, &[
        "4294967295",
        "-1",
        "4294967295",
        "4294967295",
        "4294967295",
    ]),
    test_decimal_precision_checks => (r##"
        fun main() {
            println(0.1f + 0.2f)
            println((0.1 + 0.2) == 0.3)
            println((1_000_000_000_000.0).toLong())
            println((1.0 / 0.0).isInfinite())
            println((-1.0 / 0.0).isInfinite())
        }
    "##, &[
        "0.30000001",
        "false",
        "1000000000000",
        "true",
        "true",
    ]),
    test_nan_and_infinity_parse => (r##"
        fun main() {
            val nan = 0.0 / 0.0
            val inf = 1.0 / 0.0
            val ninf = -1.0 / 0.0
            println(nan.isNaN())
            println(inf.isInfinite())
            println(ninf.isInfinite())
            println((inf + inf).isInfinite())
        }
    "##, &[
        "true",
        "true",
        "true",
        "true",
    ]),
    test_integer_to_string_roundtrip => (r##"
        fun main() {
            val value = 123
            val s = value.toString()
            println(s)
            println(s.toInt())
            println((-45).toString())
            println((-45).toString().toIntOrNull())
        }
    "##, &[
        "123",
        "123",
        "-45",
        "-45",
    ]),
    test_float_to_string_and_back => (r##"
        fun main() {
            val value = 12.5
            val text = value.toString()
            println(text)
            println(text.toDouble())
            println(2.0.toString())
            println(2.0.toString().toIntOrNull())
        }
    "##, &[
        "12.5",
        "12.5",
        "2.0",
        "null",
    ]),
    test_boolean_to_int_cast => (r##"
        fun main() {
            val a: Boolean = true
            val b: Boolean = false
            println(if (a) 1 else 0)
            println(if (b) 1 else 0)
        }
    "##, &[
        "1",
        "0",
    ]),
}
