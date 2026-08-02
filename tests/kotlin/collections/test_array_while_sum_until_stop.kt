// vybe-test: kotlin/collections/test_array_while_sum_until_stop
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun main() {
            val nums = arrayOf(3, 1, 0, 9, 2)
            var index = 0
            var sum = 0
            while (index < nums.size) {
                if (nums[index] == 0) {
                    break
                }
                sum += nums[index]
                index += 1
            }
            println(sum)
            println(index)
        }

