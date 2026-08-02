// vybe-test: kotlin/collections_maps_ops/test_map_get_or_default_for_missing_key
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("one" to 1, "two" to 2)
            __check((map.getOrDefault("one", -1)).toString(), "1")
            __check((map.getOrDefault("three", -1)).toString(), "-1")
        }
