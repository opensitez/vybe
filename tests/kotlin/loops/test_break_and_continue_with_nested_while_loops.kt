// vybe-test: kotlin/loops/test_break_and_continue_with_nested_while_loops
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
            var i = 0
            var outerTotal = 0
            while (i < 3) {
                var j = 0
                while (j < 4) {
                    j += 1
                    if (j == 2) continue
                    if (i == 1 && j == 4) break
                    outerTotal += i + j
                }
                i += 1
            }
            __p((outerTotal).toString())
        
__check("15")
}
