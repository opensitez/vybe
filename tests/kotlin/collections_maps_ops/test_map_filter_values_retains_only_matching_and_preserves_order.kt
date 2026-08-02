// vybe-test: kotlin/collections_maps_ops/test_map_filter_values_retains_only_matching_and_preserves_order
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            val filtered = map.filterValues { it >= 2 }
            __check((filtered.keys.joinToString(",")).toString(), "b,c")
            __check((filtered.values.joinToString(",")).toString(), "2,3")
        }
