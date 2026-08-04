// vybe-test: kotlin/named_arguments/test_named_arguments_named_and_default_in_chained_calls
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun base(x: Int = 1, y: Int = 2): Int = x + y
        fun scale(x: Int, y: Int = 2): Int = x * y
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
            __p((base(y = 9)).toString())
            __p((scale(3, y = 5)).toString())
        
__check("10\n15")
}
