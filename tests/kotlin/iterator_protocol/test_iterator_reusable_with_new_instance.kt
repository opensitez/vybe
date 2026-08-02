// vybe-test: kotlin/iterator_protocol/test_iterator_reusable_with_new_instance
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val src = listOf(1, 2)
            val a = src.iterator()
            val b = src.iterator()
            __check((a.toList().joinToString(",")).toString(), "1,2")
            __check((b.toList().joinToString(",")).toString(), "1,2")
        }
