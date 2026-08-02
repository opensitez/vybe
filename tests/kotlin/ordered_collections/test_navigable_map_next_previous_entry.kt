// vybe-test: kotlin/ordered_collections/test_navigable_map_next_previous_entry
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = java.util.TreeMap<Int, String>()
            map[1] = "one"
            map[3] = "three"
            map[5] = "five"
            __check((map.higherKey(3)).toString(), "5")
            __check((map.lowerKey(3)).toString(), "1")
        }
