// vybe-test: kotlin/java_util_collections/test_java_collections_index_and_last_index_of_sublist
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3, 2, 3, 4))
            val sub = java.util.ArrayList<Int>(listOf(2, 3))
            __check((java.util.Collections.indexOfSubList(values, sub)).toString(), "1")
            __check((java.util.Collections.lastIndexOfSubList(values, sub)).toString(), "3")
            __check((java.util.Collections.indexOfSubList(values, java.util.ArrayList<Int>(listOf(9, 9)))).toString(), "-1")
        }
