// vybe-test: kotlin/non_local_returns/test_multiple_returns_from_nested_lambdas
// origin: languages/kotlin/tests/kotlin/test_non_local_returns.rs

fun firstDivisible(values: List<Int>): Int {
            values.forEach {
                if (it % 3 == 0) {
                    return it
                }
            }
            return -1
        }

        fun all(values: List<Int>): Int {
            values.forEach { first ->
                if (first == 0) return 0
                if (first > 0) return@forEach
            }
            return 9
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
            __p((firstDivisible(listOf(2, 5, 6, 7))).toString())
            __p((all(listOf(-1, 1, 2))).toString())
        
__check("6\n9")
}
