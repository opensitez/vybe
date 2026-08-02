// vybe-test: kotlin/collections_maps_ops/test_map_contains_key_and_value
// origin: languages/kotlin/tests/kotlin/test_collections_maps_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 1, "b" to 2, "c" to 2)
            __check((map.containsKey("b")).toString(), "true")
            __check((map.containsValue(2)).toString(), "true")
            __check((map.containsValue(3)).toString(), "false")
        }
