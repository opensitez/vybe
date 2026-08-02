// vybe-test: kotlin/loops/test_for_with_destructuring_map_pairs
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            val values = mapOf("x" to 3, "y" to 4)
            var total = 0
            for ((key, value) in values) {
                if (key == "x") {
                    total += value
                } else {
                    total += value * 2
                }
            }
            println(total)
        }

