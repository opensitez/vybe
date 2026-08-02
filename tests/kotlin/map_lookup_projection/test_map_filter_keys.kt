// vybe-test: kotlin/map_lookup_projection/test_map_filter_keys
// origin: languages/kotlin/tests/kotlin/test_map_lookup_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mapOf("a" to 1, "bb" to 2, "ccc" to 3)
            val short = source.filterKeys { it.length <= 2 }
            __check((short.size).toString(), "2")
            __check((short.keys.joinToString(",")).toString(), "a,bb")
        }
