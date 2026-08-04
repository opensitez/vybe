// vybe-test: kotlin/loop_labels/test_label_for_do_while_style
// origin: languages/kotlin/tests/kotlin/test_loop_labels.rs

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
            var i = 0
            var out = ""
            mark@ for (ch in 1..3) {
                out += ch.toString()
                if (ch == 2) continue@mark
                out += "x"
            }
            __p((out).toString())
        
__check("1x2x3x")
}
