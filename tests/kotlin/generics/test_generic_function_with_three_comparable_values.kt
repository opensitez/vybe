// vybe-test: kotlin/generics/test_generic_function_with_three_comparable_values
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T : Comparable<T>> maxOfThree(a: T, b: T, c: T): T {
            return if (a > b && a > c) a else if (b > c) b else c
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
            __p((maxOfThree(4, 9, 1)).toString())
            __p((maxOfThree("alpha", "gamma", "beta")).toString())
        
__check("9\ngamma")
}
