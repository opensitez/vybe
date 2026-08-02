// vybe-test: kotlin/iterator_protocol/test_iterator_flat_map_concats
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun main() {
            val it = listOf(listOf(1, 2), listOf(3)).flatMap { it }.iterator()
            val values = mutableListOf<Int>()
            while (it.hasNext()) values.add(it.next())
            println(values.joinToString(","))
        }

