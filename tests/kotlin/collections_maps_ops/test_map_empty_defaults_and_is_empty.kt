// vybe-test: kotlin/collections_maps_ops/test_map_empty_defaults_and_is_empty
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = emptyMap<String, Int>()
            __check((map.isEmpty()).toString(), "true")
            __check((map.getOrDefault("x", 5)).toString(), "5")
            __check((map.orEmpty().size).toString(), "0")
        }
