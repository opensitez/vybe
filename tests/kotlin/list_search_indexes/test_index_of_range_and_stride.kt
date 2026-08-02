// vybe-test: kotlin/list_search_indexes/test_index_of_range_and_stride
// origin: languages/kotlin/tests/kotlin/test_list_search_indexes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(10, 20, 30, 40)
            __check((values.indices.step(2).joinToString(",")).toString(), "0,2")
            __check((values.slice(1..3).size).toString(), "3")
            __check((values.subList(1, 3).size).toString(), "2")
        }
