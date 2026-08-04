// vybe-test: kotlin/recursion/test_recursion_find_max
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun maxOf(values: List<Int>): Int {
            if (values.size == 1) return values[0]
            val tail = values.drop(1)
            val candidate = maxOf(tail)
            return if (values[0] > candidate) values[0] else candidate
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
            __p((maxOf(listOf(3, 1, 9, 2))).toString())
        
__check("9")
}
