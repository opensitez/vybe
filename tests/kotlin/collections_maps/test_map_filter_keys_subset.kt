// vybe-test: kotlin/collections_maps/test_map_filter_keys_subset
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val metrics = mapOf("alpha" to 1, "beta" to 2, "gamma" to 3)
            val short = metrics.filterKeys { it.length == 4 }
            __check((short.size).toString(), "2")
            __check((short["beta"]).toString(), "2")
            __check((short.containsKey("alpha")).toString(), "false")
        }
