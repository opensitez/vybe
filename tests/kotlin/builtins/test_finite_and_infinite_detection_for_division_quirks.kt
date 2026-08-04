// vybe-test: kotlin/builtins/test_finite_and_infinite_detection_for_division_quirks
// origin: languages/kotlin/tests/kotlin/test_builtins.rs

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
            __p(((1.0 / 3.0).isFinite()).toString())
            __p(((1.0 / 0.0).isInfinite()).toString())
            __p(((-1.0 / 0.0).isInfinite()).toString())
            __p(((0.0 / 0.0).isNaN()).toString())
        
__check("true\ntrue\ntrue\ntrue")
}
