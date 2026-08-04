// vybe-test: kotlin/loops/test_for_on_reversed_range_uses_expected_order_for_nested_accumulation
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
            var sequence = ""
            for (i in 3 downTo 1) {
                for (j in 0..1) {
                    sequence += i
                    sequence += ":"
                }
            }
            __p((sequence).toString())
        
__check("3:3:2:2:1:1:")
}
