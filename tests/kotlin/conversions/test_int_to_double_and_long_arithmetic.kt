// vybe-test: kotlin/conversions/test_int_to_double_and_long_arithmetic
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Int = 7
            val asDouble = value.toDouble()
            val doubled = asDouble * 2.0
            __check((doubled).toString(), "14")
            __check((doubled.toInt()).toString(), "14")
        }
