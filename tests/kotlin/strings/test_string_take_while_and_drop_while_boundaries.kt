// vybe-test: kotlin/strings/test_string_take_while_and_drop_while_boundaries
// origin: languages/kotlin/tests/kotlin/test_strings.rs

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
            val value = "12abc34"
            __p((value.takeWhile { it.isDigit() }).toString())
            __p((value.dropWhile { it.isDigit() }).toString())
            __p((value.takeLastWhile { it.isDigit() }).toString())
            __p((value.dropLastWhile { it.isDigit() }).toString())
        
__check("12\nabc34\n34\n12abc")
}
