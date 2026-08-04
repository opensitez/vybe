// vybe-test: kotlin/java_util_arrays/test_java_arrays_long_copy_of_range_and_sort
// origin: languages/kotlin/tests/kotlin/test_java_util_arrays.rs

var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = longArrayOf(9L, 4L, 7L, 1L, 8L, 2L)
            val segment = java.util.Arrays.copyOfRange(data, 1, 5)
            java.util.Arrays.sort(segment)
            __p((java.util.Arrays.toString(segment)).toString())
        
__check("[1, 4, 7, 8]")
}
