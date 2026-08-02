// vybe-test: kotlin/tuples/test_tuple_pair_zip_is_empty_when_lengths_mismatch
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val zipped = listOf(1, 2).zip(listOf("x", "y", "z"))
            __check((zipped.size).toString(), "2")
            __check((zipped[1].first).toString(), "2")
            __check((zipped[1].second).toString(), "y")
        }
