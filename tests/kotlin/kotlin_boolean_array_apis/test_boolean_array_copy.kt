// vybe-test: kotlin/kotlin_boolean_array_apis/test_boolean_array_copy
// origin: languages/kotlin/tests/kotlin/test_kotlin_boolean_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = booleanArrayOf(true, false)
            val copy = source.copyOf()
            copy[0] = false
            __check((source[0].toString()).toString(), "true")
            __check((copy[0].toString()).toString(), "false")
        }
