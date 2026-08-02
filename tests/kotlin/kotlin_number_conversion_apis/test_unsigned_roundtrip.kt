// vybe-test: kotlin/kotlin_number_conversion_apis/test_unsigned_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_number_conversion_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val u = (-1).toUInt()
            __check((u).toString(), "4294967295")
            __check((u.toInt()).toString(), "-1")
            __check((u.toLong()).toString(), "4294967295")
            val w = u.toULong()
            __check((w).toString(), "4294967295")
            __check((w.toUInt()).toString(), "4294967295")
        }
