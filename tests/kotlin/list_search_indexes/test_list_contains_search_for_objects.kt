// vybe-test: kotlin/list_search_indexes/test_list_contains_search_for_objects
// origin: languages/kotlin/tests/kotlin/test_list_search_indexes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Box(val v: Int)
            val values = listOf(Box(1), Box(2), Box(1))
            val a = Box(1)
            __check((values.contains(a)).toString(), "false")
            __check((values.indexOfFirst { it.v == 2 }).toString(), "1")
            __check((values.indexOfLast { it.v == 1 }).toString(), "2")
        }
