// vybe-test: kotlin/java_util_collections/test_java_collections_min_with_comparator
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.ArrayList<String>(listOf("aa", "z", "bbb", "c"))
            val shortest = java.util.Collections.min(values, compareBy<String> { it.length })
            val longest = java.util.Collections.max(values, compareBy<String> { it.length })
            __check((shortest).toString(), "z")
            __check((longest).toString(), "bbb")
        }
