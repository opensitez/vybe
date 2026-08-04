// vybe-test: kotlin/tailrec_functions/test_tailrec_find_first_non_zero
// origin: languages/kotlin/tests/kotlin/test_tailrec_functions.rs

tailrec fun firstNonZero(values: List<Int>, idx: Int = 0): Int {
            return if (idx >= values.size) -1 else if (values[idx] != 0) values[idx] else firstNonZero(values, idx + 1)
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
            __p((firstNonZero(listOf(0, 0, 9, 1))).toString())
        
__check("9")
}
