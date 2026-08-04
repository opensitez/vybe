// vybe-test: kotlin/scoping_functions/test_run_keeps_outer_reference_unchanged_after_block
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class State(var value: Int)

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
            val state = State(4)
            val block = state.run {
                value = value * 3
                value + 1
            }
            __p((state.value).toString())
            __p((block).toString())
        
__check("12\n13")
}
