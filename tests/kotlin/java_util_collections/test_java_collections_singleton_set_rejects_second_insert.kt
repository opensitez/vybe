// vybe-test: kotlin/java_util_collections/test_java_collections_singleton_set_rejects_second_insert
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun main() {
            val values: java.util.Set<String> = java.util.Collections.singleton("x")
            println(values.size)
            try {
                values.add("y")
                println("added")
            } catch (ex: Exception) {
                println("error")
            }
        }

