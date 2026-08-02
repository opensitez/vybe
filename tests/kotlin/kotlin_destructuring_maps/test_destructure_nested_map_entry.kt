// vybe-test: kotlin/kotlin_destructuring_maps/test_destructure_nested_map_entry
// origin: languages/kotlin/tests/kotlin/test_kotlin_destructuring_maps.rs

fun main() {
            val groups = mapOf("x" to mapOf("a" to 1), "y" to mapOf("b" to 2))
            var total = 0
            for ((outer, inner) in groups) {
                for ((innerKey, innerValue) in inner) {
                    total += innerValue
                }
            }
            println(total)
        }

