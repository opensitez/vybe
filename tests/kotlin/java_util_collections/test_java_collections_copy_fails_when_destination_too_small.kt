// vybe-test: kotlin/java_util_collections/test_java_collections_copy_fails_when_destination_too_small
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun main() {
            val source = java.util.ArrayList<Int>(listOf(1, 2, 3))
            val target = java.util.ArrayList<Int>(listOf(0))
            try {
                java.util.Collections.copy(target, source)
                println("ok")
            } catch (ex: Exception) {
                println("error")
            }
        }

