// vybe-test: kotlin/conversions/test_float_conversion_roundtrip_candidate
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 3.75f
            val asString = value.toString()
            val asDouble = asString.toDouble()
            __check((asString).toString(), "3.75")
            __check((asDouble).toString(), "3.75")
            __check((asDouble == 3.75).toString(), "true")
        }
