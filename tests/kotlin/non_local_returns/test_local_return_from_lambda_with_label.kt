// vybe-test: kotlin/non_local_returns/test_local_return_from_lambda_with_label
// origin: languages/kotlin/tests/kotlin/test_non_local_returns.rs

fun sumEven(values: List<Int>): Int {
            var total = 0
            values.forEach {
                if (it % 2 == 0) return@forEach
                total += it
            }
            return total
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
            __p((sumEven(listOf(1, 2, 3, 4, 5))).toString())
        
__check("9")
}
