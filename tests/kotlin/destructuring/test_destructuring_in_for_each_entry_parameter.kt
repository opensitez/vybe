// vybe-test: kotlin/destructuring/test_destructuring_in_for_each_entry_parameter
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

fun main() {
            val pairs = listOf(Pair(1, 2), Pair(3, 4))
            var total = 0
            pairs.forEach { (left, right) ->
                total += left + right
            }
            println(total)
        }

