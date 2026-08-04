// vybe-test: kotlin/non_local_returns/test_nested_non_local_return_with_for_loop
// origin: languages/kotlin/tests/kotlin/test_non_local_returns.rs

fun firstOdd(values: List<Int>): Int {
            values.forEach {
                run {
                    if (it % 2 == 1) {
                        return it
                    }
                }
            }
            return -1
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
            __p((firstOdd(listOf(2, 4, 6, 9, 10))).toString())
        
__check("9")
}
