// vybe-test: kotlin/tuples/test_tuple_triple_function_with_return_position
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun stats(): Triple<Int, Int, Int> {
                return Triple(2, 4, 6)
            }
            val score = stats()
            __check((score.third / score.second).toString(), "1")
            __check((score.first + score.second + score.third).toString(), "12")
        }
