// vybe-test: kotlin/map_lookup_projection/test_map_filter_pairs_combined
// origin: languages/kotlin/tests/kotlin/test_map_lookup_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mapOf("aa" to 1, "bb" to 2, "ac" to 3)
            val projected = source.filter { it.key.startsWith("a") && it.value > 1 }
            __check((projected.size).toString(), "1")
            __check((projected["ac"]).toString(), "3")
        }
