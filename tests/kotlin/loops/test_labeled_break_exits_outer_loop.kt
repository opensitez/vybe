// vybe-test: kotlin/loops/test_labeled_break_exits_outer_loop
// origin: languages/kotlin/tests/kotlin/test_loops.rs

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
            var total = 0
            outer@ for (i in 1..5) {
                for (j in 1..5) {
                    total += i + j
                    if (i == 3 && j == 2) break@outer
                }
            }
            __p((total).toString())
        
__check("21")
}
