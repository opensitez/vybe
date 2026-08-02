// vybe-test: kotlin/iterator_protocol/test_iterable_for_each_index
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun main() {
            val values = listOf(1, 2, 3)
            val out = StringBuilder()
            for ((i, v) in values.withIndex()) {
                out.append(i).append(":").append(v).append("|")
            }
            println(out.toString())
        }

