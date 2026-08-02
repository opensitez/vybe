// vybe-test: kotlin/primitive_array_apis/test_double_array_is_not_empty
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = doubleArrayOf()
            __check((values.isEmpty()).toString(), "true")
            __check((values.isNotEmpty()).toString(), "false")
        }
