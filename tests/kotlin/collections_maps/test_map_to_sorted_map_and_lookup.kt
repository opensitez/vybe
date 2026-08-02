// vybe-test: kotlin/collections_maps/test_map_to_sorted_map_and_lookup
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val metrics = mapOf(3 to "c", 1 to "a", 2 to "b")
            val sorted = metrics.toSortedMap()
            __check((sorted.keys.first()).toString(), "1")
            __check((sorted.keys.last()).toString(), "3")
            __check((sorted[2]).toString(), "b")
        }
