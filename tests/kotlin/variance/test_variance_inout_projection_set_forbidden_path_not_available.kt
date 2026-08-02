// vybe-test: kotlin/variance/test_variance_inout_projection_set_forbidden_path_not_available
// origin: languages/kotlin/tests/kotlin/test_variance.rs

fun accepts(outItems: List<out String>) {
            __check((outItems.size).toString(), "3")
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            accepts(listOf("a", "b", "c"))
        }
