// vybe-test: kotlin/tuples/test_tuple_pair_values_replace_in_mutable_collection
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pairs = mutableListOf(Pair(1, 2), Pair(3, 4))
            pairs[1] = Pair(5, 6)
            __check((pairs[0]).toString(), "(1, 2)")
            __check((pairs[1]).toString(), "(5, 6)")
            __check((pairs.size).toString(), "2")
        }
