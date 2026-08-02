// vybe-test: kotlin/map_lookup_projection/test_map_keys_projection
// origin: languages/kotlin/tests/kotlin/test_map_lookup_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mapOf("a" to 1, "bb" to 2)
            val upper = source.mapKeys { it.key.uppercase() }
            __check((upper.keys.joinToString(",")).toString(), "A,BB")
            __check((upper["A"]).toString(), "1")
        }
