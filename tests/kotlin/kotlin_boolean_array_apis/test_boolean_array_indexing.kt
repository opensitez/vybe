// vybe-test: kotlin/kotlin_boolean_array_apis/test_boolean_array_indexing
// origin: languages/kotlin/tests/kotlin/test_kotlin_boolean_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = booleanArrayOf(true, false, true)
            __check((data[0].toString()).toString(), "true")
            __check((data[2].toString()).toString(), "true")
        }
