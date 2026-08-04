// vybe-test: kotlin/java_util_arrays/test_java_arrays_as_list_search_and_size
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
            val data = arrayOf("a", "b", "c")
            val view = java.util.Arrays.asList(data)
            __p((view.size).toString())
            __p((view.indexOf("b")).toString())
            __p((view.contains("c")).toString())
        
__check("3\n1\ntrue")
}
