// vybe-test: kotlin/operators/test_nested_elvis_chain_and_rhs_evaluation_boundary
// origin: languages/kotlin/tests/kotlin/test_operators.rs

var fallbackCalls = 0

        fun fallback(value: String?): String {
            fallbackCalls += 1
            return value ?: "default"
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
            val first: String? = null
            val second: String? = null
            val third: String? = "value"
            val present: String? = "keep"
            __p((first ?: second ?: fallback(third)).toString())
            __p((fallbackCalls).toString())
            fallbackCalls = 0
            __p((present ?: fallback(present)).toString())
            __p((fallbackCalls).toString())
        
__check("value\n1\nkeep\n0")
}
