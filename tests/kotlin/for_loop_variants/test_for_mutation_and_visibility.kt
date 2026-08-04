// vybe-test: kotlin/for_loop_variants/test_for_mutation_and_visibility
// origin: languages/kotlin/tests/kotlin/test_for_loop_variants.rs

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
            var values = intArrayOf(1, 2, 3)
            for (i in values.indices) {
                values[i] = values[i] * 2
            }
            __p((values[0] + values[1] + values[2]).toString())
        
__check("12")
}
