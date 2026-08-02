// vybe-test: kotlin/collections_maps/test_list_iteration_with_manual_early_exit
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val values = listOf(4, 8, 12, 16)
            var sum = 0
            var found = false
            for (value in values) {
                sum += value
                if (sum > 15) {
                    found = true
                    break
                }
            }
            println(sum)
            println(found)
        }

