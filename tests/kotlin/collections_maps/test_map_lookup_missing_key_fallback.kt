// vybe-test: kotlin/collections_maps/test_map_lookup_missing_key_fallback
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val scores = mapOf("a" to 1, "b" to 2)
            __check((scores["missing"] ?: -1).toString(), "-1")
            __check((scores.containsKey("missing")).toString(), "false")
            __check((scores.get("b") ?: -1).toString(), "2")
        }
