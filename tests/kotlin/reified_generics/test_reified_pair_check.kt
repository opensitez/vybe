// vybe-test: kotlin/reified_generics/test_reified_pair_check
// origin: languages/kotlin/tests/kotlin/test_reified_generics.rs

inline fun <reified T> isPair(value: Any?): String = if (value is Pair<T, T>) "pair" else "not-pair"

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((isPair<Int>(Pair(1, 2))).toString(), "pair")
            __check((isPair<String>(Pair(1, "x"))).toString(), "not-pair")
        }
