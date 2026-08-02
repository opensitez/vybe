// vybe-test: kotlin/kotlin_number_conversion_apis/test_boolean_to_int_cast
// origin: languages/kotlin/tests/kotlin/test_kotlin_number_conversion_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a: Boolean = true
            val b: Boolean = false
            __check((if (a) 1 else 0).toString(), "1")
            __check((if (b) 1 else 0).toString(), "0")
        }
