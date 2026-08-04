// vybe-test: kotlin/kotlin_return_expressions/test_returning_result_from_nested_block
// origin: languages/kotlin/tests/kotlin/test_kotlin_return_expressions.rs

fun compute(v: Int): Int {
            val out = run {
                if (v < 5) {
                    return@run v * 2
                }
                v
            }
            return out
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
            __p((compute(3)).toString())
            __p((compute(7)).toString())
        
__check("6\n7")
}
