// vybe-test: kotlin/tuples/test_tuple_pair_zip_default_no_transform_is_pair
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = listOf(1, 2, 3)
            val right = listOf("a", "b", "c")
            val zipped = left.zip(right)
            __check((zipped.size).toString(), "3")
            __check((zipped[0]).toString(), "(1, a)")
            __check((zipped[1].first + zipped[1].second.length).toString(), "3")
        }
