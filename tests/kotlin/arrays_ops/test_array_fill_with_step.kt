// vybe-test: kotlin/arrays_ops/test_array_fill_with_step
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun main() {
            val nums = IntArray(6)
            for (i in nums.indices step 2) {
                nums[i] = i
            }
            println(nums.joinToString(","))
        }

