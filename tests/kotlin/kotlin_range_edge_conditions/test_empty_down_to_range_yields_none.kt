// vybe-test: kotlin/kotlin_range_edge_conditions/test_empty_down_to_range_yields_none
// origin: languages/kotlin/tests/kotlin/test_kotlin_range_edge_conditions.rs

fun main() {
            var total = 0
            for (i in 3 downTo 7) {
                total += i
            }
            println(total)
        }

