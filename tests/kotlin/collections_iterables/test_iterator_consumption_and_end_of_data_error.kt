// vybe-test: kotlin/collections_iterables/test_iterator_consumption_and_end_of_data_error
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun main() {
            val iterator = listOf(1, 2).iterator()
            println(iterator.hasNext())
            println(iterator.next())
            println(iterator.hasNext())
            println(iterator.next())
            println(iterator.hasNext())
            try {
                iterator.next()
                println("should_not_happen")
            } catch (e: NoSuchElementException) {
                println("error")
            }
        }

