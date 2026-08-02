// vybe-test: kotlin/kotlin_star_projections/test_star_projection_count
// origin: languages/kotlin/tests/kotlin/test_kotlin_star_projections.rs

fun isPresent(values: List<*>): String {
            return if (values.isNotEmpty()) "has" else "empty"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((isPresent(listOf(1))).toString(), "has")
            __check((isPresent(listOf<Int?>())).toString(), "empty")
        }
