// vybe-test: kotlin/destructuring/test_destructure_from_array_of_pairs
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun main() {
            var first = 0
            var second = 0
            for (cell in arrayOf(Pair(3, 4), Pair(1, 2))) {
                val (x, y) = cell
                first += x
                second += y
            }
            println(first)
            println(second)
        }

