// vybe-test: kotlin/collections_set/test_set_ordered_iteration_on_linked_set_preserved_after_mutation
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun main() {
            val values = linkedSetOf(2, 1)
            values.add(3)
            values.remove(1)
            values.add(1)
            var order = ""
            for (value in values) {
                order += value.toString()
            }
            println(order)
            println(values.size)
        }

