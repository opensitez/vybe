// vybe-test: kotlin/kotlin_number_conversion_apis/test_number_to_int_from_float
// origin: languages/kotlin/tests/kotlin/test_kotlin_number_conversion_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((3.9.toInt()).toString(), "3")
            __check(((-3.9).toInt()).toString(), "-3")
            __check((3.4f.toInt()).toString(), "3")
            __check(((-3.4f).toInt()).toString(), "-3")
        }
