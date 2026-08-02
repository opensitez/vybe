// vybe-test: kotlin/java_util_collections/test_java_collections_swap
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3))
            java.util.Collections.swap(values, 0, 2)
            __check((values).toString(), "[3, 2, 1]")
        }
