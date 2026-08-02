// vybe-test: kotlin/kotlin_star_projections/test_star_projection_first
// origin: languages/kotlin/tests/kotlin/test_kotlin_star_projections.rs

fun firstValue(items: List<*>): String {
            if (items.isEmpty()) {
                return "none"
            }
            return items[0].toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val mixed: List<Any?> = listOf("x", 2, true)
            __check((firstValue(mixed)).toString(), "x")
            __check((firstValue(listOf())).toString(), "none")
        }
