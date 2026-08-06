// vybe-test: kotlin/bitwise_operations/test_bitwise_counting_subset_flags
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

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
            val values = listOf(0b1010, 0b1111, 0b1000, 0b0011)
            val countAnyHigh = values.count { it and 0b1000 != 0 }
            val countZeroLow = values.count { it and 1 == 0 }
            val countPairs = values.filter { (it and 0b0110) == 0b0010 }
            __p((countAnyHigh).toString())
            __p((countZeroLow).toString())
            __p((countPairs.joinToString(",")).toString())
        
// Real Kotlin agrees: infix `and` binds tighter than `==`, so
// `it and 1 == 0` keeps only the even values 0b1010 and 0b1000 (2), and
// the pair filter passes BOTH 0b1010 (10) and 0b0011 (3).
__check("3\n2\n10,3")
}
