// vybe-test: kotlin/collections_iterables/test_list_with_index_iteration_binding
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun main() {
            val letters = listOf("a", "b", "c")
            var parts = ""
            for ((i, value) in letters.withIndex()) {
                parts += "${i}:${value};"
            }
            println(parts)
        }

