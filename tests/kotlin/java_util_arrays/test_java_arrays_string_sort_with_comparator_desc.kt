// vybe-test: kotlin/java_util_arrays/test_java_arrays_string_sort_with_comparator_desc
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = arrayOf("aa", "b", "cccc", "ddd")
            java.util.Arrays.sort(data, java.util.Comparator { a, b ->
                b.length - a.length
            })
            __check((java.util.Arrays.toString(data)).toString(), "[cccc, ddd, aa, b]")
        }
