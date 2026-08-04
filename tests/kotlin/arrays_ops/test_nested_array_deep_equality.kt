// vybe-test: kotlin/arrays_ops/test_nested_array_deep_equality
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

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
            val a = arrayOf(arrayOf(1, 2), arrayOf(3))
            val b = arrayOf(arrayOf(1, 2), arrayOf(3))
            val c = arrayOf(arrayOf(1, 2), arrayOf(4))
            __p((a.contentDeepEquals(b)).toString())
            __p((a.contentDeepEquals(c)).toString())
            __p((a.contentDeepToString()).toString())
        
__check("true\nfalse\n[[1, 2], [3]]")
}
