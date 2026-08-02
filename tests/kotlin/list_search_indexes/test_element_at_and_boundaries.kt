// vybe-test: kotlin/list_search_indexes/test_element_at_and_boundaries
// origin: languages/kotlin/tests/kotlin/test_list_search_indexes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(11, 22, 33)
            __check((values.elementAt(2)).toString(), "33")
            __check((values.elementAtOrNull(5) ?: "none").toString(), "none")
            __check((values.elementAtOrElse(5) { value -> value * 10 + 1 }).toString(), "51")
        }
