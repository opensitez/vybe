// vybe-test: kotlin/iterator_protocol/test_iterable_sequence_chain
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val it = generateSequence(1) { it + 2 }
                .take(4)
                .toList()
                .iterator()
            __check((it.next()).toString(), "1")
            __check((it.next()).toString(), "3")
            __check((it.next()).toString(), "5")
            __check((it.next()).toString(), "7")
            __check((it.hasNext()).toString(), "false")
        }
