// vybe-test: kotlin/collections_maps_ops/test_map_map_values_then_filter_keys_chain
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 1, "bb" to 2, "ccc" to 3)
            val result = map
                .mapValues { it.value * 2 }
                .filterKeys { it.length <= 2 }
            __check((result["a"]).toString(), "2")
            __check((result["bb"]).toString(), "4")
            __check((result["ccc"] ?: -1).toString(), "-1")
        }
