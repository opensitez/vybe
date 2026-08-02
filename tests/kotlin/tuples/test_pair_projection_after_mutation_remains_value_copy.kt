// vybe-test: kotlin/tuples/test_pair_projection_after_mutation_remains_value_copy
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = mutableListOf(1, 2)
            val (left, right) = Pair(source[0], source[1])
            source[0] = 9
            __check((left).toString(), "1")
            __check((right).toString(), "2")
        }
