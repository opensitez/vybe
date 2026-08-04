// vybe-test: kotlin/bitwise_operations/test_bitwise_roundtrip_with_mask
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
            val original = 0b10101010
            val mask = 0b11110000
            val hidden = original and mask
            val shown = original and mask.inv()
            val visible = (original and mask.inv())
            __p((hidden).toString())
            __p((shown).toString())
            __p((visible).toString())
            __p((hidden + visible).toString())
        
__check("160\n10\n10\n170")
}
