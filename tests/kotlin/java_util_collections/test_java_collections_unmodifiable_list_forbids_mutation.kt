// vybe-test: kotlin/java_util_collections/test_java_collections_unmodifiable_list_forbids_mutation
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3))
            val safe = java.util.Collections.unmodifiableList(values)
            println(safe[1])
            try {
                safe[1] = 9
                println("changed")
            } catch (ex: Exception) {
                println("error")
            }
        }

