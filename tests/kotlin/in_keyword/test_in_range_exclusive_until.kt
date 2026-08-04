// vybe-test: kotlin/in_keyword/test_in_range_exclusive_until
// origin: languages/kotlin/tests/kotlin/test_in_keyword.rs

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
            __p((5 in 1 until 5).toString())
            __p((4 in 1 until 5).toString())
            __p((0 until 5).toString())
            __p((3 in 1..4).toString())
        
__check("false\ntrue\n1,2,3,4\ntrue")
}
