// vybe-test: kotlin/tuples/test_tuple_pair_in_when_expression
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pair = Pair("blue", 3)
            val label = when (pair) {
                Pair("red", 1) -> "red-one"
                Pair("blue", 3) -> "blue-three"
                else -> "other"
            }
            __check((label).toString(), "blue-three")
        }
