// vybe-test: kotlin/tuples/test_tuple_pair_in_function_parameter
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun sum(pair: Pair<Int, Int>): Int {
                return pair.first + pair.second
            }
            __check((sum(Pair(4, 6))).toString(), "10")
            __check((sum(8 to 1)).toString(), "9")
        }
