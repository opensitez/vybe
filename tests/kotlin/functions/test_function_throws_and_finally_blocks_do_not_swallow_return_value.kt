// vybe-test: kotlin/functions/test_function_throws_and_finally_blocks_do_not_swallow_return_value
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun risky(value: Int): Int {
            try {
                if (value < 0) {
                    throw Exception("bad")
                }
                return value * 2
            } finally {
                __p(("final").toString())
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
            __p((risky(4)).toString())
            try {
                risky(-1)
            } catch (e: Exception) {
                __p(("caught").toString())
            }
        
__check("final\n8\nfinal\ncaught")
}
