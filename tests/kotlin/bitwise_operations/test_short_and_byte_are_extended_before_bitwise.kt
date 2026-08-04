// vybe-test: kotlin/bitwise_operations/test_short_and_byte_are_extended_before_bitwise
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
            val signedByte: Byte = -1
            val signedShort: Short = -2
            val byteUnsigned = signedByte.toInt() and 0xFF
            val shortUnsigned = signedShort.toInt() and 0xFFFF
            val combined = (byteUnsigned and shortUnsigned)
            __p((byteUnsigned).toString())
            __p((shortUnsigned).toString())
            __p((combined).toString())
        
__check("255\n65534\n254")
}
