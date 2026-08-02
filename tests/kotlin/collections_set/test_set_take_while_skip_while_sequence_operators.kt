// vybe-test: kotlin/collections_set/test_set_take_while_skip_while_sequence_operators
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = linkedSetOf(1, 2, 3, 4, 5)
            __check((values.takeWhile { it < 4 }).toString(), "[1, 2, 3]")
            __check((values.dropWhile { it < 4 }).toString(), "[4, 5]")
        }
