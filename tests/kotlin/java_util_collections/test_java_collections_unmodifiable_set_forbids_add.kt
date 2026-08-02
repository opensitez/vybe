// vybe-test: kotlin/java_util_collections/test_java_collections_unmodifiable_set_forbids_add
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun main() {
            val source = java.util.HashSet<Int>()
            source.add(1)
            source.add(2)
            val safe = java.util.Collections.unmodifiableSet(source)
            try {
                safe.add(3)
                println("added")
            } catch (ex: Exception) {
                println("error")
            }
            println(safe.size)
        }

