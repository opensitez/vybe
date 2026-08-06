// vybe-test: kotlin/bitwise_operations/test_isolate_least_significant_set_bit
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
            val value = 0b1011000
            val lsb = value and (-value)
            __p((lsb).toString())
            __p(((value and (value - 1))).toString())
        
// Real Kotlin agrees: `value and (value - 1)` CLEARS the lowest set bit —
// 0b1011000 & 0b1010111 = 0b1010000 = 80.
__check("8\n80")
}
