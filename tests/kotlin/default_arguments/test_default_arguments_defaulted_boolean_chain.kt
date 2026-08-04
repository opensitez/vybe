// vybe-test: kotlin/default_arguments/test_default_arguments_defaulted_boolean_chain
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun flags(a: Boolean = true, b: Boolean = false): String = if (a && !b) "on" else "off"
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
            __p((flags()).toString())
            __p((flags(a = false)).toString())
            __p((flags(b = true)).toString())
        
__check("on\noff\noff")
}
