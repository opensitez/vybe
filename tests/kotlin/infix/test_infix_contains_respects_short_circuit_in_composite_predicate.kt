// vybe-test: kotlin/infix/test_infix_contains_respects_short_circuit_in_composite_predicate
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class Gate {
            var probes = 0
            operator fun contains(value: Int): Boolean {
                probes += 1
                return value % 2 == 0
            }
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
            val gate = Gate()
            val outcome = (1 in gate) || (2 in gate)
            __p((outcome).toString())
            __p((gate.probes).toString())
        
__check("true\n2")
}
