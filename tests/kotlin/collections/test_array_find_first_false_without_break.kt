// vybe-test: kotlin/collections/test_array_find_first_false_without_break
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun main() {
            val nums = arrayOf(1, 2, 3, 4)
            var missing = true
            for (value in nums) {
                if (value == 5) {
                    missing = false
                }
            }
            println(missing)
        }

