// vybe-test: kotlin/variance/test_variance_star_projection_read
// origin: languages/kotlin/tests/kotlin/test_variance.rs

fun read(items: List<*>, idx: Int): String {
            return items[idx]?.toString() ?: "nil"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((read(listOf("x", "y"), 0)).toString(), "x")
            __check((read(listOf<Int>(1, 2), 1)).toString(), "2")
        }
