// vybe-test: kotlin/tuples/test_tuple_pair_in_set_uniqueness
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seen = setOf(Pair(1, 2), Pair(1, 2), Pair(2, 1))
            __check((seen.size).toString(), "2")
            __check((seen.contains(Pair(2, 1))).toString(), "true")
            __check((seen.contains(Pair(3, 4))).toString(), "false")
        }
