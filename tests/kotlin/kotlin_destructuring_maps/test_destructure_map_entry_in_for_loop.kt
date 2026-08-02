// vybe-test: kotlin/kotlin_destructuring_maps/test_destructure_map_entry_in_for_loop
// origin: languages/kotlin/tests/kotlin/test_kotlin_destructuring_maps.rs

fun main() {
            val values = mapOf("x" to 1, "y" to 2)
            var total = 0
            for ((k, v) in values) {
                if (k == "x") {
                    total = v
                }
            }
            println(total)
        }

