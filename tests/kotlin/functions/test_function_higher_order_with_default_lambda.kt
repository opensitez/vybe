// vybe-test: kotlin/functions/test_function_higher_order_with_default_lambda
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun apply(value: Int, op: (Int) -> Int = { it + 1 }): Int {
            return op(value)
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
            __p((apply(4)).toString())
            __p((apply(4) { it * 3 }).toString())
        
__check("5\n12")
}
