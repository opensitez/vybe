// vybe-test: kotlin/collections_maps_ops/test_map_nullable_key_lookup_missing
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map: Map<Int?, String> = mapOf(null to "nil", 1 to "one")
            __check((map[2] ?: "none").toString(), "none")
            __check((map[null]).toString(), "nil")
            __check((map.containsKey(2)).toString(), "false")
        }
