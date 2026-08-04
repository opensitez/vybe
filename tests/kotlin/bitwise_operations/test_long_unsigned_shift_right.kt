// vybe-test: kotlin/bitwise_operations/test_long_unsigned_shift_right
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
            val negative: Long = -1L
            val signed: Long = -16L
            __p((negative ushr 1).toString())
            __p((signed ushr 2).toString())
            __p((15L ushr 1).toString())
            __p((1L ushr 1).toString())
        
__check("9223372036854775807\n2305843009213693950\n7\n0")
}
