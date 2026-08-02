// vybe-test: kotlin/tuples/test_tuple_pair_and_triple_mixed_collection_projection
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val mixed = listOf(
                Pair("p", 1),
                Triple("t", 2, 3)
            )
            __check((mixed[0]).toString(), "(p, 1)")
            __check(((mixed[1] as Triple<String, Int, Int>).second).toString(), "2")
        }
