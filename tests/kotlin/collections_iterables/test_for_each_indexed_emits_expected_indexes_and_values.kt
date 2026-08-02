// vybe-test: kotlin/collections_iterables/test_for_each_indexed_emits_expected_indexes_and_values
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun main() {
            val items = listOf("x", "y", "z")
            var marker = ""
            items.forEachIndexed { index, value ->
                marker += "${index}${value}"
            }
            println(marker)
        }

