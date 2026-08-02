// vybe-test: kotlin/variance/test_variance_inout_projection_get
// origin: languages/kotlin/tests/kotlin/test_variance.rs

fun firstOrNull(items: List<out Any>): String = items.firstOrNull()?.toString() ?: "none"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((firstOrNull(listOf(1, 2))).toString(), "1")
            __check((firstOrNull(listOf("x"))).toString(), "x")
        }
