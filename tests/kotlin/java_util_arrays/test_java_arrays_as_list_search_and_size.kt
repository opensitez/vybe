// vybe-test: kotlin/java_util_arrays/test_java_arrays_as_list_search_and_size
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = arrayOf("a", "b", "c")
            val view = java.util.Arrays.asList(data)
            __check((view.size).toString(), "3")
            __check((view.indexOf("b")).toString(), "1")
            __check((view.contains("c")).toString(), "true")
        }
