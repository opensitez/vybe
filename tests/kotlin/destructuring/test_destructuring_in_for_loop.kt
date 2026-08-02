// vybe-test: kotlin/destructuring/test_destructuring_in_for_loop
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun main() {
            var total = 0
            for (entry in arrayOf(Pair(1, 2), Pair(3, 4), Pair(5, 6))) {
                val (x, y) = entry
                total += x
                total += y
            }
            println(total)
        }

