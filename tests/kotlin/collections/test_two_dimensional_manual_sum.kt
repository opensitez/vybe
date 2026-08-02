// vybe-test: kotlin/collections/test_two_dimensional_manual_sum
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun main() {
            val grid = arrayOf(
                arrayOf(1, 2, 3),
                arrayOf(4, 5, 6)
            )
            var total = 0
            for (row in grid) {
                for (cell in row) {
                    total += cell
                }
            }
            println(total)
        }

