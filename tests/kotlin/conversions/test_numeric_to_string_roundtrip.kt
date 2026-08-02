// vybe-test: kotlin/conversions/test_numeric_to_string_roundtrip
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n = 15
            val fromNumber = n.toString()
            val parsed = fromNumber.toInt()
            __check((fromNumber).toString(), "15")
            __check((parsed).toString(), "15")
            __check((parsed == n).toString(), "true")
        }
