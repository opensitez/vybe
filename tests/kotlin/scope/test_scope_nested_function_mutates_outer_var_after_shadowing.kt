// vybe-test: kotlin/scope/test_scope_nested_function_mutates_outer_var_after_shadowing
// origin: languages/kotlin/tests/kotlin/test_scope.rs

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
            var value = 5

            fun bump(delta: Int) {
                fun total(): Int {
                    return value + delta
                }
                value = total()
            }

            bump(3)
            __p((value).toString())

            val value = 10
            fun useShadowed(): Int {
                return value + 1
            }
            __p((useShadowed()).toString())
        
__check("8\n11")
}
