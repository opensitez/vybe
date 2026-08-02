// vybe-test: kotlin/mutable_set_apis/test_mutable_set_clear_empty
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2, 3)
            values.clear()
            __check((values.isEmpty()).toString(), "true")
            __check((values.size).toString(), "0")
        }
