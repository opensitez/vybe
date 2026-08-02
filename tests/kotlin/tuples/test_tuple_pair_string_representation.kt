// vybe-test: kotlin/tuples/test_tuple_pair_string_representation
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pair = Pair("a", "b")
            __check((pair).toString(), "(a, b)")
        }
