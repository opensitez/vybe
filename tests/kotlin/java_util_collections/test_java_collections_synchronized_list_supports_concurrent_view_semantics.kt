// vybe-test: kotlin/java_util_collections/test_java_collections_synchronized_list_supports_concurrent_view_semantics
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3))
            val sync = java.util.Collections.synchronizedList(values)
            sync.add(4)
            __check((sync.size).toString(), "4")
            sync[0] = 0
            __check((sync[0]).toString(), "0")
        }
