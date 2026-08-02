// vybe-test: kotlin/iterator_protocol/test_iterator_can_consume_generator_with_sequence
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = generateSequence(1) { it + 1 }.take(3)
            val it = seq.iterator()
            __check((it.next()).toString(), "1")
            __check((it.next()).toString(), "2")
            __check((it.next()).toString(), "3")
            __check((it.hasNext()).toString(), "false")
        }
