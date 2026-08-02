// vybe-test: kotlin/arrays_ops/test_int_array_mutation_with_for_each_indexed
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun main() {
            val nums = intArrayOf(1, 2, 3)
            nums.forEachIndexed { index, value ->
                nums[index] = value * 3
            }
            println(nums.joinToString(","))
        }

