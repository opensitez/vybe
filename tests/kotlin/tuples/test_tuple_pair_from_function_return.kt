// vybe-test: kotlin/tuples/test_tuple_pair_from_function_return
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun make(): Pair<Int, Int> {
                return Pair(7, 11)
            }
            val (left, right) = make()
            __check((left + right).toString(), "18")
        }
