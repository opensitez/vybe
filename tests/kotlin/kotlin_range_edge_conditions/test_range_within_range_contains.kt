// vybe-test: kotlin/kotlin_range_edge_conditions/test_range_within_range_contains
// origin: languages/kotlin/tests/kotlin/test_kotlin_range_edge_conditions.rs

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
            val outer = 1..20
            val inner = 5..8
            __p((inner.first() in outer).toString())
            __p((inner.last() in outer).toString())
            __p((21 in outer).toString())
        
__check("true\ntrue\nfalse")
}
