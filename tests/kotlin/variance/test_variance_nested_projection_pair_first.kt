// vybe-test: kotlin/variance/test_variance_nested_projection_pair_first
// origin: languages/kotlin/tests/kotlin/test_variance.rs

val pairFirst: (Pair<out String, *>) -> String = { p -> p.first }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pairFirst(Pair("x", 3))).toString(), "x")
        }
