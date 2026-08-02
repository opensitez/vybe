// vybe-test: kotlin/collections_maps_ops/test_map_keys_and_values_views
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 3)
            __check((map.keys.joinToString(",")).toString(), "a,b,c")
            __check((map.values.sum()).toString(), "6")
            __check((map.values.maxOrNull() ?: 0).toString(), "3")
        }
