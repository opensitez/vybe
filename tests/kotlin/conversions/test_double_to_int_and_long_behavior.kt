// vybe-test: kotlin/conversions/test_double_to_int_and_long_behavior
// origin: languages/kotlin/tests/kotlin/test_conversions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = 12.9
            __check((source.toInt()).toString(), "12")
            __check((source.toLong()).toString(), "12")
        }
