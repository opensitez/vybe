// vybe-test: kotlin/ranges/test_reversed_range_iteration_order_and_bounds
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun main() {
            val forward = (1..7).reversed()
            var forwardFirst = ""
            for (value in forward) {
                forwardFirst += value.toString()
            }
            val reversed = (7 downTo 1).reversed()
            var reversedFirst = ""
            for (value in reversed) {
                reversedFirst += value.toString()
            }
            println(forward.first())
            println(forward.last())
            println(reversedFirst)
        }

