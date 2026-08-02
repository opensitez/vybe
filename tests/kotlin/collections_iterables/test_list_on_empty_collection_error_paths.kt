// vybe-test: kotlin/collections_iterables/test_list_on_empty_collection_error_paths
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun main() {
            try {
                val x = emptyList<Int>().first()
                println(x)
            } catch (e: NoSuchElementException) {
                println("first_error")
            }
            try {
                val y = emptyList<Int>().elementAt(1)
                println(y)
            } catch (e: IndexOutOfBoundsException) {
                println("element_error")
            }
        }

