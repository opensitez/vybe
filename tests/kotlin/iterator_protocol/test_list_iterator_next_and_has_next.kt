// vybe-test: kotlin/iterator_protocol/test_list_iterator_next_and_has_next
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val it = listOf(1, 2, 3).iterator()
            val b1 = it.hasNext()
            val v1 = it.next()
            val b2 = it.hasNext()
            val v2 = it.next()
            __check((b1).toString(), "true")
            __check((v1).toString(), "1")
            __check((b2).toString(), "true")
            __check((v2).toString(), "2")
            __check((it.next()).toString(), "3")
        }
