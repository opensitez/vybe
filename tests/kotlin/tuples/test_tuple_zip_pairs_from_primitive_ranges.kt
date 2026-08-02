// vybe-test: kotlin/tuples/test_tuple_zip_pairs_from_primitive_ranges
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val zipped = (1..4).zip(10 downTo 7)
            __check((zipped.size).toString(), "4")
            __check((zipped[0]).toString(), "(1, 10)")
            __check((zipped[3]).toString(), "(4, 7)")
        }
