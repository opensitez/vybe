// vybe-test: kotlin/scope_shadowing/test_nested_if_shadowing_isolated
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

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
            val n = 10
            fun test(x: Int): Int {
                val n = x + 1
                return if (x > 5) {
                    val n = n + 5
                    n
                } else {
                    n
                }
            }
            __p((test(6)).toString())
            __p((n).toString())
        
__check("12\n10")
}
