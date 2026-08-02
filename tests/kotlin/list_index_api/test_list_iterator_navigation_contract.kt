// vybe-test: kotlin/list_index_api/test_list_iterator_navigation_contract
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val it = listOf(1, 2, 3).listIterator()
            __check((it.hasNext()).toString(), "true")
            __check((it.next()).toString(), "1")
            __check((it.next()).toString(), "2")
            __check((it.hasPrevious()).toString(), "true")
            __check((it.previous()).toString(), "1")
            __check((it.previousIndex()).toString(), "0")
        }
