// vybe-test: kotlin/function_overloads/test_overload_with_no_argument_and_named_defaults
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun flag(v: Int = 1): String = "n" + v
        fun flag(v: String = "s"): String = "s" + v
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
            __p((flag()).toString())
            __p((flag(2)).toString())
            __p((flag(v = "x")).toString())
        
__check("n1\nn2\nsx")
}
