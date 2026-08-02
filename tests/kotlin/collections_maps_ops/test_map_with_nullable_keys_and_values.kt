// vybe-test: kotlin/collections_maps_ops/test_map_with_nullable_keys_and_values
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map: Map<String?, Int?> = mapOf(null to 5, "x" to null)
            __check((map[null]).toString(), "5")
            __check((map.containsKey(null)).toString(), "true")
            __check((map["x"] ?: -1).toString(), "-1")
            __check((map.getOrElse(null) { -1 }).toString(), "5")
        }
