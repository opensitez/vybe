// vybe-test: kotlin/arrays_ops/test_int_array_index_of_and_last_index_of
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
            val nums = intArrayOf(4, 1, 4, 2, 4)
            __p((nums.indexOf(4)).toString())
            __p((nums.lastIndexOf(4)).toString())
            __p((nums.indexOf(9)).toString())
        
__check("0\n4\n-1")
}
