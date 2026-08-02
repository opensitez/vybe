// vybe-test: kotlin/iterator_protocol/test_iterator_drop_while_to_list
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = generateSequence(0) { it + 1 }
                .take(6)
                .dropWhile { it < 3 }
                .toList()
            __check((values.joinToString(",")).toString(), "3,4,5")
        }
