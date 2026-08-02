// vybe-test: kotlin/collections_maps_ops/test_map_sorted_with_custom_comparator
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 3)
            val reversed = map.toSortedMap(compareByDescending { it })
            __check((reversed.keys.joinToString(",")).toString(), "c,b,a")
            __check((reversed.values.joinToString(",")).toString(), "3,2,1")
        }
