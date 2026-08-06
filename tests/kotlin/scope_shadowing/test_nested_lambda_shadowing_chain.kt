// vybe-test: kotlin/scope_shadowing/test_nested_lambda_shadowing_chain
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
            val prefix = "A"
            val f = { prefix: String ->
                { prefix: Int -> "${prefix}_${prefix + 1}" }
            }
            val g = f("B")
            __p((g(3)).toString())
            __p((prefix).toString())
        
__check("3_4\nA")
}
