// vybe-test: kotlin/iterator_protocol/test_iterator_over_map_keys
// origin: languages/kotlin/tests/kotlin/test_iterator_protocol.rs

fun main() {
            val map = linkedMapOf("a" to 1, "b" to 2)
            val values = StringBuilder()
            for (k in map.keys) {
                values.append(k)
            }
            println(values.toString())
            println(map.keys.count())
        }

