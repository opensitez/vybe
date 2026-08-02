// vybe-test: kotlin/collections/test_array_range_for_indices_bounds
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun main() {
            val nums = arrayOf(1, 2, 3, 4, 5)
            var evenTotal = 0
            for (i in nums.indices) {
                if (i % 2 == 1) evenTotal += nums[i]
            }
            println(evenTotal)
            println(nums.lastIndex)
        }

