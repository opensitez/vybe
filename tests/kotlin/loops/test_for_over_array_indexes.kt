// vybe-test: kotlin/loops/test_for_over_array_indexes
// origin: languages/kotlin/tests/kotlin/test_loops.rs

fun main() {
            val nums = arrayOf(10, 20, 30)
            var sum = 0
            for (i in nums.indices) {
                sum += i + nums[i]
            }
            println(sum)
        }

