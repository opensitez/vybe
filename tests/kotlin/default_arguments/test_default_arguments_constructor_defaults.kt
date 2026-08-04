// vybe-test: kotlin/default_arguments/test_default_arguments_constructor_defaults
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

class Box(val value: Int = 1, val label: String = "x")
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
            val a = Box()
            val b = Box(2)
            val c = Box(label = "z")
            __p((a.value).toString())
            __p((a.label).toString())
            __p((b.value).toString())
            __p((c.label).toString())
        
__check("1\nx\n2\nz")
}
