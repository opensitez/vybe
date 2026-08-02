// vybe-test: kotlin/list_index_api/test_list_sublist_slices_view
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2, 3, 4, 5)
            val part = values.subList(1, 4)
            part[0] = 9
            __check((values.joinToString(",")).toString(), "1,9,3,4,5")
            __check((part.joinToString(",")).toString(), "9,3,4")
        }
