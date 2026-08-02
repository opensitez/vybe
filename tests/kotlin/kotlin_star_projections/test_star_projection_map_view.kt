// vybe-test: kotlin/kotlin_star_projections/test_star_projection_map_view
// origin: languages/kotlin/tests/kotlin/test_kotlin_star_projections.rs

fun firstKey(values: Map<*, *>): String {
            if (values.isEmpty()) {
                return "none"
            }
            return values.keys.iterator().next().toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((firstKey(mapOf("a" to 1, "b" to 2))).toString(), "a")
            __check((firstKey(mapOf<String, Int>())).toString(), "none")
        }
