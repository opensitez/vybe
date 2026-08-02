// vybe-test: kotlin/generics/test_generic_map_projection_of_collections
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> toStringMap(values: Map<String, T>): Map<String, String> {
            return values.mapValues { it.value.toString() }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val input: Map<String, Int> = mapOf("a" to 1, "b" to 2)
            val projected = toStringMap(input)
            __check((projected["a"]).toString(), "1")
            __check((projected["b"]).toString(), "2")
        }
