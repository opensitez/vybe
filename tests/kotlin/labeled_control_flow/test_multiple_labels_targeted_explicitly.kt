// vybe-test: kotlin/labeled_control_flow/test_multiple_labels_targeted_explicitly
// origin: languages/kotlin/tests/kotlin/test_labeled_control_flow.rs

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
            var x = 0
            outer@ for (i in 0..1) {
                inner@ for (j in 0..2) {
                    if (i == 1 && j == 0) {
                        continue@outer
                    }
                    if (j == 2) {
                        break@inner
                    }
                    x += 1
                }
            }
            __p((x).toString())
        
__check("4")
}
