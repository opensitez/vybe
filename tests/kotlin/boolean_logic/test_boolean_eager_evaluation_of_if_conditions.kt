// vybe-test: kotlin/boolean_logic/test_boolean_eager_evaluation_of_if_conditions
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

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
            val a = true
            val b = false
            val c = if (a && b) "both" else "not both"
            val d = if (a || b) "some" else "none"
            __p((c).toString())
            __p((d).toString())
        
__check("not both\nsome")
}
