// vybe-test: kotlin/java_util_collections/test_java_collections_unmodifiable_map_forbids_put
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun main() {
            val source = java.util.LinkedHashMap<String, Int>()
            source["a"] = 1
            source["b"] = 2
            val safe = java.util.Collections.unmodifiableMap(source)
            try {
                safe["c"] = 3
                println("added")
            } catch (ex: Exception) {
                println("error")
            }
            println(safe["b"])
        }

