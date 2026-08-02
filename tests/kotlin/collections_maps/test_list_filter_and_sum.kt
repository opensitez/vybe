// vybe-test: kotlin/collections_maps/test_list_filter_and_sum
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            val evens = nums.filter { it % 2 == 0 }
            var total = 0
            for (v in evens) {
                total += v
            }
            println(evens.size)
            println(total)
        }

