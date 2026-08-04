// vybe-test: kotlin/bitwise_operations/test_bitwise_with_java_long_bitcount_and_number_of_leading_zeros
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
            val sample = 0b0001_0010
            __p((java.lang.Integer.bitCount(sample)).toString())
            __p((java.lang.Integer.numberOfLeadingZeros(sample)).toString())
            __p((java.lang.Integer.numberOfTrailingZeros(sample)).toString())
            __p((java.lang.Integer.numberOfTrailingZeros(0)).toString())
        
__check("2\n28\n1\n32")
}
