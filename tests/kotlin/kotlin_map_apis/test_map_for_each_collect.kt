// vybe-test: kotlin/kotlin_map_apis/test_map_for_each_collect
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun main() {
            val map = linkedMapOf("a" to 3, "b" to 5)
            var total = 0
            map.forEach { _, value -> total += value }
            println(total)
        }

