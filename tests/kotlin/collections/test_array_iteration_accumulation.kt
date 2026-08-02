// vybe-test: kotlin/collections/test_array_iteration_accumulation
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun main() {
            val nums = arrayOf(1, 2, 3)
            var sum = 0
            for (n in nums) {
                sum += n
            }
            println(sum)
        }

