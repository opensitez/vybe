// vybe-test: kotlin/mutable_list_apis/test_mutable_list_sort_descending
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(4, 1, 3, 2)
            values.sortDescending()
            __check((values.joinToString(",")).toString(), "4,3,2,1")
        }
