// vybe-test: kotlin/java_util_collections/test_java_collections_disjoint_when_no_shared_elements
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = java.util.Arrays.asList(1, 2, 3)
            val b = java.util.Arrays.asList(4, 5, 6)
            __check((java.util.Collections.disjoint(a, b)).toString(), "true")
            __check((java.util.Collections.disjoint(a, java.util.Arrays.asList(3, 8))).toString(), "false")
        }
