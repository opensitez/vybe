// vybe-test: kotlin/kotlin_collection_factories/test_build_list_can_skip_by_condition
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun main() {
            val values = buildList {
                for (v in 1..6) {
                    if (v % 2 == 0) add(v)
                }
            }
            println(values.joinToString(","))
        }

