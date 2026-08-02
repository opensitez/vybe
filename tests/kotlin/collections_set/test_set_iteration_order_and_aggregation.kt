// vybe-test: kotlin/collections_set/test_set_iteration_order_and_aggregation
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun main() {
            val values = linkedSetOf(3, 1, 2, 4)
            var output = ""
            for (value in values) {
                output += value.toString()
            }
            println(output)
            println(values.first())
            println(values.elementAt(2))
        }

