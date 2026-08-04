// vybe-test: kotlin/kotlin_sorting_comparators/test_partition_by_parity
// origin: languages/kotlin/tests/kotlin/test_kotlin_sorting_comparators.rs

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
            val (even, odd) = listOf(1, 2, 3, 4, 5).partition { it % 2 == 0 }
            __p((even.joinToString(",")).toString())
            __p((odd.joinToString(",")).toString())
        
__check("2,4\n1,3,5")
}
