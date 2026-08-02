// vybe-test: kotlin/list_search_indexes/test_find_index_by_contains_check
// origin: languages/kotlin/tests/kotlin/test_list_search_indexes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("kotlin", "java", "python")
            val found = values.indexOfFirst { it.startsWith("jav") }
            val missing = values.indexOfFirst { it.startsWith("rust") }
            __check((found).toString(), "1")
            __check((missing).toString(), "-1")
        }
