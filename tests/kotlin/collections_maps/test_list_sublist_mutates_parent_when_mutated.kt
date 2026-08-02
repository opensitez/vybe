// vybe-test: kotlin/collections_maps/test_list_sublist_mutates_parent_when_mutated
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2, 3, 4, 5)
            val window = values.subList(1, 4)
            window[1] = 30
            window.removeAt(2)
            __check((values.joinToString(",")).toString(), "1,2,30,5")
            __check((window.size).toString(), "2")
        }
