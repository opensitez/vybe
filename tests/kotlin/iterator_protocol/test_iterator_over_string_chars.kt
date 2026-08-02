// vybe-test: kotlin/iterator_protocol/test_iterator_over_string_chars
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val it = "ab".iterator()
            var first = it.next()
            var second = it.next()
            __check((first).toString(), "a")
            __check((second).toString(), "b")
            __check((it.hasNext()).toString(), "false")
        }
