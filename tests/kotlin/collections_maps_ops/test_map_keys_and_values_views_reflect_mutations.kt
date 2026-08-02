// vybe-test: kotlin/collections_maps_ops/test_map_keys_and_values_views_reflect_mutations
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mutableMapOf("a" to 1, "b" to 2)
            val keys = map.keys
            val values = map.values
            map["a"] = 4
            map["c"] = 3
            __check((keys.contains("c")).toString(), "true")
            __check((values.any { it == 4 }).toString(), "true")
            __check((values.sum()).toString(), "9")
        }
