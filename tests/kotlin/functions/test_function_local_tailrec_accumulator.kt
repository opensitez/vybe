// vybe-test: kotlin/functions/test_function_local_tailrec_accumulator
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun power(base: Int, exp: Int): Int {
            tailrec fun loop(remaining: Int, acc: Int): Int {
                if (remaining == 0) return acc
                return loop(remaining - 1, acc * base)
            }
            return loop(exp, 1)
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
            __p((power(2, 0)).toString())
            __p((power(3, 3)).toString())
        
__check("1\n27")
}
