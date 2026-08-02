// vybe-test: kotlin/java_util_collections/test_java_collections_add_all_appends_elements
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2))
            val more = java.util.ArrayList<Int>(listOf(3, 4))
            val changed = java.util.Collections.addAll(values, more[0], more[1])
            __check((changed).toString(), "true")
            __check((values).toString(), "[1, 2, 3, 4]")
        }
