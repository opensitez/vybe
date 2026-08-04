// vybe-test: kotlin/spread_arguments/test_spread_multiple_arrays_to_vararg
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun join(values: IntArray): Int {
            var total = 0
            for (v in values) total += v
            return total
        }
        fun sum(prefix: String, vararg values: Int): Int {
            return prefix.length + values.sum()
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
            val a = intArrayOf(1, 2)
            val b = intArrayOf(3, 4)
            __p((sum("x", *a, *b)).toString())
        
__check("9")
}
