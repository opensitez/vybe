// vybe-test: kotlin/iterator_protocol/test_iterator_for_each_contract
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun main() {
            val it = listOf("a", "b").iterator()
            it.forEachRemaining { println(it) }
        }

