// vybe-test: kotlin/collections_iterables/test_iterator_reuse_on_iterable_is_distinct_instances
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun main() {
            val values = listOf(1, 2, 3)
            val first = values.iterator()
            while (first.hasNext()) {
                println(first.next())
            }
            val second = values.iterator()
            println(second.hasNext())
            println(second.next())
        }

