// vybe-test: kotlin/collections/test_array_range_filter_without_library
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun main() {
            val nums = arrayOf(1, 2, 3, 4, 5, 6)
            var odds = ""
            for (i in nums.indices) {
                if (i % 2 == 1) {
                    odds += nums[i].toString()
                }
            }
            println(odds)
        }

