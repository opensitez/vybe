// vybe-test: kotlin/ranges/test_range_bounds_are_evaluated_once
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

var leftBoundCalls = 0
        var rightBoundCalls = 0

        fun left(): Int {
            leftBoundCalls += 1
            return 1
        }

        fun right(): Int {
            rightBoundCalls += 1
            return 4
        }

        fun main() {
            var total = 0
            for (value in left()..right()) {
                total += value
            }
            println(leftBoundCalls)
            println(rightBoundCalls)
            println(total)
        }

