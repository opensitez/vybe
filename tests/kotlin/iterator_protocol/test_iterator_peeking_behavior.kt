// vybe-test: kotlin/iterator_protocol/test_iterator_peeking_behavior
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

class P : Iterator<Int> {
            private val data = listOf(7, 8)
            private var index = 0
            override fun hasNext() = index < data.size
            override fun next() = data[index++]
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val it = P()
            __check((it.hasNext()).toString(), "true")
            __check((it.next()).toString(), "7")
            __check((it.hasNext()).toString(), "true")
            __check((it.next()).toString(), "8")
            __check((it.hasNext()).toString(), "false")
        }
