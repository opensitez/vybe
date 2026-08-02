// vybe-test: kotlin/variance/test_variance_nested_projection_pair_second_readonly
// origin: languages/kotlin/tests/kotlin/test_variance.rs

val pairSecond: (Pair<*, out Number>) -> String = { p -> p.second.toString() }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pairSecond(Pair("x", 9L))).toString(), "9")
            __check((pairSecond(Pair(1, 4.5))).toString(), "4.5")
        }
