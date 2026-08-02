// vybe-test: kotlin/collections/test_array_find_by_linear_scan
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun main() {
            val nums = arrayOf(7, 3, 9, 4)
            var found = -1
            var index = 0
            for (value in nums) {
                if (value == 9) {
                    found = value
                    break
                }
                index += 1
            }
            println(found)
            println(index)
        }

