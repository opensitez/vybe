// vybe-test: kotlin/arrays_ops/test_array_out_of_bounds_raises
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun main() {
            val nums = intArrayOf(1, 2, 3)
            try {
                println(nums[3])
            } catch (e: IndexOutOfBoundsException) {
                println("out_of_bounds")
            }
        }

