// vybe-test: kotlin/ordered_collections/test_map_keys_view_is_ordered_for_linked
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf(1 to "a", 2 to "b", 3 to "c")
            val view = map.keys
            __check((view.joinToString(",")).toString(), "1,2,3")
            map.remove(2)
            map[4] = "d"
            __check((view.joinToString(",")).toString(), "1,3,4")
        }
