// vybe-test: kotlin/ranges/test_reversed_range_iteration_order_and_bounds
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
            val forward = (1..7).reversed()
            var forwardFirst = ""
            for (value in forward) {
                forwardFirst += value.toString()
            }
            val reversed = (7 downTo 1).reversed()
            var reversedFirst = ""
            for (value in reversed) {
                reversedFirst += value.toString()
            }
            __p((forward.first()).toString())
            __p((forward.last()).toString())
            __p((reversedFirst).toString())
        
__check("7\n1\n1234567")
}
