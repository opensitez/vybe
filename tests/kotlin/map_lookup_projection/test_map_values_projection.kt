// vybe-test: kotlin/map_lookup_projection/test_map_values_projection
// origin: languages/kotlin/tests/kotlin/test_map_lookup_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mapOf("a" to 1, "b" to 2)
            val doubled = source.mapValues { it.value * 2 }
            __check((doubled.values.joinToString(",")).toString(), "2,4")
            __check((doubled["b"]).toString(), "4")
        }
