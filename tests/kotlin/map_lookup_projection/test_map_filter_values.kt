// vybe-test: kotlin/map_lookup_projection/test_map_filter_values
// origin: languages/kotlin/tests/kotlin/test_map_lookup_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mapOf("x" to 3, "y" to 12, "z" to 18)
            val filtered = source.filterValues { it % 6 == 0 }
            __check((filtered.keys.joinToString(",")).toString(), "y,z")
            __check((filtered["z"]).toString(), "18")
        }
