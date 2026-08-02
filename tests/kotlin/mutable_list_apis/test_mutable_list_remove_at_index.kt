// vybe-test: kotlin/mutable_list_apis/test_mutable_list_remove_at_index
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(7, 8, 9)
            val removed = values.removeAt(1)
            __check((removed).toString(), "8")
            __check((values.joinToString(",")).toString(), "7,9")
        }
