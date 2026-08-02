// vybe-test: kotlin/collections_maps/test_list_sublist_mutates_parent_when_cleared
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf("a", "b", "c", "d")
            val window = values.subList(1, 3)
            window.clear()
            __check((values.joinToString(",")).toString(), "a,d")
            __check((window.size).toString(), "0")
        }
