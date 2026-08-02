// vybe-test: kotlin/control_flow/test_for_down_to_step_negative_offset
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var nums = ""
            for (n in 9 downTo 3 step 2) {
                nums += n.toString()
            }
            println(nums)
        }

