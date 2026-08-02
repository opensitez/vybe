// vybe-test: kotlin/iterator_protocol/test_iterator_drop_while_like_manual
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun main() {
            val it = listOf(1, 2, 3, 4).iterator()
            while (it.hasNext()) {
                val n = it.next()
                if (n < 3) continue
                print(n)
            }
            println("")
        }

