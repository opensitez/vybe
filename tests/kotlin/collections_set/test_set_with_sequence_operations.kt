// vybe-test: kotlin/collections_set/test_set_with_sequence_operations
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1, 2, 3, 4, 5)
            val sequence = values.asSequence()
            val total = sequence.filter { it > 2 }.sum()
            __check((total).toString(), "12")
        }
