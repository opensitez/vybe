// vybe-test: kotlin/iterator_protocol/test_iterator_yield_sum_reduce
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun main() {
            val it = (1..4).iterator()
            var total = 0
            while (it.hasNext()) {
                total += it.next()
            }
            println(total)
        }

