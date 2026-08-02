// vybe-test: kotlin/mutable_list_apis/test_mutable_list_index_of_and_last_index_of
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2, 3, 2)
            __check((values.indexOf(2)).toString(), "1")
            __check((values.lastIndexOf(2)).toString(), "3")
        }
