// vybe-test: kotlin/ranges/test_long_range_iteration
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

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
            var total = 0L
            for (value in 1L..7L step 2) {
                total += value
            }
            __p((total).toString())
            __p((6L in 1L..7L).toString())
            __p((8L in 1L until 8L).toString())
        
__check("16\ntrue\nfalse")
}
