// vybe-test: kotlin/tuples/test_pair_destructuring_captures_expected_values
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (left, right) = Pair(8, 13)
            __check((left).toString(), "8")
            __check((right).toString(), "13")
        }
