// vybe-test: kotlin/range_projection/test_range_down_to_and_reversed
// origin: languages/kotlin/tests/kotlin/test_range_projection.rs

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
            val r = 10 downTo 4
            __p((r.toList().joinToString(",")).toString())
            __p((r.first).toString())
            __p((r.last).toString())
            val asc = r.reversed()
            __p((asc.toList().joinToString(",")).toString())
        
__check("10,9,8,7,6,5,4\n10\n4\n4,5,6,7,8,9,10")
}
