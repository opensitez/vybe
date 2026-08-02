// vybe-test: kotlin/operators/test_range_and_contains_for_custom_window
// origin: languages/kotlin/tests/kotlin/test_operators.rs

class Window(val low: Int, val high: Int) {
            operator fun contains(value: Int): Boolean {
                return value >= low && value <= high
            }

            operator fun rangeTo(other: Int): IntRange {
                return low..other
            }
        }

        fun main() {
            val window = Window(1, 4)
            println(2 in window)
            println(6 in window)
            val span = window..5
            var total = 0
            for (value in span) {
                total += value
            }
            println(total)
        }

