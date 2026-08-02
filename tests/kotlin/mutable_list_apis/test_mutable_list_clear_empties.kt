// vybe-test: kotlin/mutable_list_apis/test_mutable_list_clear_empties
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2)
            values.clear()
            __check((values.isEmpty()).toString(), "true")
            __check((values.size).toString(), "0")
        }
