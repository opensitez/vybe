// vybe-test: kotlin/collections_maps/test_map_filter_to_map_and_keys_set
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val scores = mapOf("a" to 1, "b" to 4, "c" to 2, "d" to 5)
            val high = scores.filterValues { it >= 4 }
            __check((high.size).toString(), "2")
            __check((high.keys.joinToString(",")).toString(), "b,d")
            __check((high.values.joinToString(",")).toString(), "4,5")
        }
