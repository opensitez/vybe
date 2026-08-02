// vybe-test: kotlin/tuples/test_tuple_pair_equality_and_inequality
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Pair(1, 2) == Pair(1, 2)).toString(), "true")
            __check((Pair(1, 2) == Pair(2, 1)).toString(), "false")
            __check((Pair(1, 2) != Pair(2, 1)).toString(), "true")
        }
