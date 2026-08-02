// vybe-test: kotlin/function_overloads/test_overload_with_pair_inputs
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun shape(v: Pair<Int, Int>): String = "pair"
        fun shape(v: Pair<String, String>): String = "sPair"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((shape(Pair(1, 2))).toString(), "pair")
            __check((shape(Pair("a", "b"))).toString(), "sPair")
        }
