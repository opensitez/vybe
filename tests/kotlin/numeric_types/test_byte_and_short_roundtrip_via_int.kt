// vybe-test: kotlin/numeric_types/test_byte_and_short_roundtrip_via_int
// origin: languages/kotlin/tests/kotlin/test_numeric_types.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b: Byte = 127
            val s: Short = 32767
            __check((b.toInt() + 1).toString(), "128")
            __check((s.toInt() + 1).toString(), "32768")
            __check((b.toLong() - 7).toString(), "120")
            __check((s.toLong() - 7).toString(), "32760")
        }
