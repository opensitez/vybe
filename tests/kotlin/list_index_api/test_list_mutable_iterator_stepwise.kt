// vybe-test: kotlin/list_index_api/test_list_mutable_iterator_stepwise
// origin: languages/kotlin/tests/kotlin/test_list_index_api.rs

fun main() {
            val iterator = mutableListOf(5, 6, 7, 8).iterator()
            var total = 0
            while (iterator.hasNext()) {
                total += iterator.next()
            }
            println(total)
            println(iterator.hasNext())
        }

