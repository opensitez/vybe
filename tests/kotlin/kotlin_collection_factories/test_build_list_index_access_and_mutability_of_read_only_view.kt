// vybe-test: kotlin/kotlin_collection_factories/test_build_list_index_access_and_mutability_of_read_only_view
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun main() {
            val values = buildList(4) {
                for (i in 0 until 4) {
                    add(i * 3)
                }
            }
            println(values[0])
            println(values[2])
            println(values.size)
        }

