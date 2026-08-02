// vybe-test: kotlin/tuples/test_tuple_infix_to_constructs_pair
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pair = "k" to 9
            __check((pair.first).toString(), "k")
            __check((pair.second).toString(), "9")
        }
