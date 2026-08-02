// vybe-test: kotlin/loops/test_for_over_array_elements_uses_current_values
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            val nums = intArrayOf(2, 4, 6)
            var total = 0
            for (n in nums) {
                total += n
            }
            println(total)
        }

