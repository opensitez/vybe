// vybe-test: kotlin/bitwise_operations/test_masking_chain_preserves_expected_bits
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
            val value = 0b10101111
            val lowNibble = value and 0x0F
            val upperNibble = (value and 0xF0) ushr 4
            __p((lowNibble).toString())
            __p((upperNibble).toString())
            __p((((upperNibble shl 4) or lowNibble)).toString())
        
__check("15\n10\n175")
}
