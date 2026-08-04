// vybe-test: kotlin/control_flow/test_if_expression_skips_false_branch_side_effects
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

var hits = 0

        fun bump(): Int {
            hits += 1
            return 0
        }

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
            val value = if (1 == 1) {
                7
            } else {
                bump()
            }
            __p((value).toString())
            __p((hits).toString())
        
__check("7\n0")
}
