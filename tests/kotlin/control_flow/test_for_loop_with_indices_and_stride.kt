// vybe-test: kotlin/control_flow/test_for_loop_with_indices_and_stride
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            val nums = arrayOf(2, 4, 6, 8, 10)
            var total = 0
            for (i in nums.indices step 2) {
                total += nums[i]
            }
            println(total)
        }

