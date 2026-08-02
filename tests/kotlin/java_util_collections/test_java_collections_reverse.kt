// vybe-test: kotlin/java_util_collections/test_java_collections_reverse
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3))
            java.util.Collections.reverse(values)
            __check((values).toString(), "[3, 2, 1]")
        }
