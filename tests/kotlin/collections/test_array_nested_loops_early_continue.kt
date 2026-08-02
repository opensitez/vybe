// vybe-test: kotlin/collections/test_array_nested_loops_early_continue
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun main() {
            val grid = arrayOf(
                arrayOf(1, 2, 3),
                arrayOf(4, 0, 6),
                arrayOf(7, 8, 9)
            )
            var total = 0
            for (row in grid) {
                for (value in row) {
                    if (value == 0) {
                        continue
                    }
                    total += value
                }
            }
            println(total)
        }

