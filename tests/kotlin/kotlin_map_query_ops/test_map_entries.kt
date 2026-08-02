// vybe-test: kotlin/kotlin_map_query_ops/test_map_entries
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_query_ops.rs

fun main() {
            val m = mapOf("a" to 1, "b" to 2)
            var sum = 0
            for ((k, v) in m.entries) {
                if (k == "a") sum += v
            }
            println(sum.toString())
            println(m.entries.size)
        }

