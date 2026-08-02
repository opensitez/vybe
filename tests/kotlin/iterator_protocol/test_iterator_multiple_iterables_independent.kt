// vybe-test: kotlin/iterator_protocol/test_iterator_multiple_iterables_independent
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(1, 2)
            val a = list.iterator()
            val b = list.iterator()
            __check((a.next()).toString(), "1")
            __check((b.next()).toString(), "1")
            __check((a.next()).toString(), "2")
            __check((b.next()).toString(), "2")
        }
