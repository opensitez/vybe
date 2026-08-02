// vybe-test: kotlin/iterator_protocol/test_iterator_over_set_is_unique
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun main() {
            val seen = linkedSetOf(1, 2, 3)
            val out = StringBuilder()
            for (v in seen) {
                out.append(v)
            }
            println(out.toString())
        }

