// vybe-test: kotlin/local_functions/test_local_function_nested_across_scopes
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

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
            fun outer(v: Int): Int {
                fun inner1(x: Int): Int = x + 1
                if (v > 0) {
                    fun inner2(y: Int): Int = inner1(y * 2)
                    return inner2(v)
                }
                return inner1(v)
            }
            __p((outer(3)).toString())
            __p((outer(0)).toString())
        
__check("7\n1")
}
