// vybe-test: kotlin/extension_functions/test_extension_function_on_pair_exposes_projection_like_accessor
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun Pair<Int, Int>.delta(): Int = second - first

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Pair(4, 9)
            __check((value.delta()).toString(), "5")
        }
