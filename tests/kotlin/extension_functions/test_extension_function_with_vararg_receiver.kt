// vybe-test: kotlin/extension_functions/test_extension_function_with_vararg_receiver
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun Int.joinWith(vararg values: Int): String {
            var total = this
            for (value in values) {
                total += value
            }
            return total.toString()
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
            __p((1.joinWith(2, 3, 4)).toString())
            __p((0.joinWith()).toString())
        
__check("10\n0")
}
