// vybe-test: kotlin/collections/test_array_last_index_mutation_pattern
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun main() {
            val nums = arrayOf(2, 4, 6, 8)
            nums[nums.lastIndex] = nums[0] + nums[1]
            var output = 0
            for (i in 0 until nums.size) {
                output += nums[i]
            }
            println(nums[3])
            println(output)
        }

