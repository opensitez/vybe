// vybe-test: kotlin/collections/test_array_with_indices_loop
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun main() {
            val nums = arrayOf(4, 9, 2)
            var output = 0
            for (i in nums.indices) {
                output += nums[i] * i
            }
            println(output)
        }

