// vybe-test: kotlin/collections_maps/test_list_sublist_invalid_range_is_runtime_error
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

fun main() {
            val values = mutableListOf(1, 2, 3)
            try {
                values.subList(3, 1)
                println("no-error")
            } catch (e: Exception) {
                println("error")
            }
        }

