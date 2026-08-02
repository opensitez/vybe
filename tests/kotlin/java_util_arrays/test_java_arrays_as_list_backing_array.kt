// vybe-test: kotlin/java_util_arrays/test_java_arrays_as_list_backing_array
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = arrayOf("one", "two", "three")
            val view = java.util.Arrays.asList(data)
            view[1] = "changed"
            __check((data[1]).toString(), "changed")
            __check((view.joinToString(",")).toString(), "one,changed,three")
        }
