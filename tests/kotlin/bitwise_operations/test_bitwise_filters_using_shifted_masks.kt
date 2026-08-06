// vybe-test: kotlin/bitwise_operations/test_bitwise_filters_using_shifted_masks
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
            val values = listOf(0, 1, 2, 3, 4, 5, 6, 7, 8, 15, 16, 31)
            val maskedTwoBits = values.map { it and 0b11 }
            val flags = values.filter { (it and 0b1000) == 0b1000 }
            __p((maskedTwoBits.joinToString(",")).toString())
            __p((flags.joinToString(",")).toString())
        
// Real Kotlin agrees: 31 = 0b11111 has bit 3 set, so the flags filter
// keeps 8, 15 AND 31.
__check("0,1,2,3,0,1,2,3,0,3,0,3\n8,15,31")
}
