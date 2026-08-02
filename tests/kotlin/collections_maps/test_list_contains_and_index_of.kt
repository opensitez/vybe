// vybe-test: kotlin/collections_maps/test_list_contains_and_index_of
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("a", "b", "c")
            __check((words.contains("b")).toString(), "true")
            __check((words.indexOf("c")).toString(), "2")
            __check((words.lastIndex).toString(), "2")
        }
