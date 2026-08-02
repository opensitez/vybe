// vybe-test: kotlin/iterator_protocol/test_iterator_filter_then_to_set
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val filtered = (1..5).asSequence().iterator().asSequence().filter { it % 2 == 0 }.toSet()
            __check((filtered.joinToString(",")).toString(), "2,4")
        }
