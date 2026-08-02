// vybe-test: kotlin/kotlin_boolean_array_apis/test_boolean_array_equals_reference
// origin: languages/kotlin/tests/kotlin/test_kotlin_boolean_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = booleanArrayOf(true, false)
            val b = booleanArrayOf(true, false)
            __check(((a == b).toString()).toString(), "false")
        }
