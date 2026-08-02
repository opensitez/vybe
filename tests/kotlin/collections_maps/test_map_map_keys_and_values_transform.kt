// vybe-test: kotlin/collections_maps/test_map_map_keys_and_values_transform
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val input = mapOf("a" to 1, "b" to 2)
            val keys = input.mapKeys { it.key.uppercase() }
            val values = input.mapValues { it.value * 10 }
            __check((keys["A"]).toString(), "1")
            __check((values["b"]).toString(), "20")
            __check((keys.size).toString(), "2")
        }
