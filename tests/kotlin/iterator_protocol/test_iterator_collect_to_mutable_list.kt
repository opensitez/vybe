// vybe-test: kotlin/iterator_protocol/test_iterator_collect_to_mutable_list
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun main() {
            val src = listOf(4, 5, 6).iterator()
            val dst = mutableListOf<Int>()
            while (src.hasNext()) {
                dst.add(src.next())
            }
            println(dst.joinToString(","))
        }

