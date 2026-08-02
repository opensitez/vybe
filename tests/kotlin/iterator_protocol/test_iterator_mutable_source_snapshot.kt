// vybe-test: kotlin/iterator_protocol/test_iterator_mutable_source_snapshot
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val src = mutableListOf(1, 2, 3)
            val it = src.iterator()
            __check((it.next()).toString(), "1")
            src.add(4)
            __check((it.next()).toString(), "2")
            __check((it.next()).toString(), "3")
        }
