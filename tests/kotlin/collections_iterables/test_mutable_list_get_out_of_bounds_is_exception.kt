// vybe-test: kotlin/collections_iterables/test_mutable_list_get_out_of_bounds_is_exception
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun main() {
            val nums = mutableListOf(10)
            try {
                println(nums[3])
            } catch (e: Exception) {
                println("oob")
            }
        }

