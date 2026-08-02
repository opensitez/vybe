// vybe-test: kotlin/variance/test_variance_star_projection_write_not_attempted
// origin: languages/kotlin/tests/kotlin/test_variance.rs

fun count(items: MutableList<*>) : Int = items.size
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((count(mutableListOf(1, 2, 3))).toString(), "3")
        }
