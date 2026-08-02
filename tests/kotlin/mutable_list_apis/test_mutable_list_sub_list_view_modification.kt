// vybe-test: kotlin/mutable_list_apis/test_mutable_list_sub_list_view_modification
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2, 3, 4)
            val window = values.subList(1, 3)
            window.clear()
            __check((values.joinToString(",")).toString(), "1,4")
        }
