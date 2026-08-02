// vybe-test: kotlin/mutable_list_apis/test_mutable_list_remove_first
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(9, 1, 2)
            __check((values.removeFirst()).toString(), "9")
            __check((values.joinToString(",")).toString(), "1,2")
        }
