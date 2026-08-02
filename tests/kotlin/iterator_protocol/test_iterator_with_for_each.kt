// vybe-test: kotlin/iterator_protocol/test_iterator_with_for_each
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun main() {
            var acc = 0
            for (v in listOf(1, 2, 3)) {
                acc += v
            }
            println(acc)
        }

