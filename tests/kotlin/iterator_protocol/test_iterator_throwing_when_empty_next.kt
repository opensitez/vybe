// vybe-test: kotlin/iterator_protocol/test_iterator_throwing_when_empty_next
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun main() {
            val it = emptyList<Int>().iterator()
            try {
                it.next()
                println("no")
            } catch (e: NoSuchElementException) {
                println("error")
            }
        }

