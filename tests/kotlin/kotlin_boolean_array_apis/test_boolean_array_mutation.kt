// vybe-test: kotlin/kotlin_boolean_array_apis/test_boolean_array_mutation
// origin: languages/kotlin/tests/kotlin/test_kotlin_boolean_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = booleanArrayOf(true, false)
            data[1] = true
            __check((data[1].toString()).toString(), "true")
        }
