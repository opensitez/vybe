// vybe-test: kotlin/map_lookup_projection/test_map_to_pair_lists_round_trip
// origin: languages/kotlin/tests/kotlin/test_map_lookup_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mapOf("x" to 9, "y" to 8)
            val pairs = source.toList()
            val rebuilt = pairs.toMap()
            __check((pairs.joinToString("|") { it.toString() }).toString(), "(x, 9)|(y, 8)")
            __check((rebuilt.size).toString(), "2")
            __check((rebuilt["y"]).toString(), "8")
        }
