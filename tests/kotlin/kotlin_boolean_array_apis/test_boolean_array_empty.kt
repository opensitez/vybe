// vybe-test: kotlin/kotlin_boolean_array_apis/test_boolean_array_empty
// origin: languages/kotlin/tests/kotlin/test_kotlin_boolean_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val empty = booleanArrayOf()
            __check((empty.size).toString(), "0")
            __check((empty.isEmpty().toString()).toString(), "true")
        }
