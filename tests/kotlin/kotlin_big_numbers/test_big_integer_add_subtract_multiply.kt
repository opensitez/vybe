// vybe-test: kotlin/kotlin_big_numbers/test_big_integer_add_subtract_multiply
// origin: languages/kotlin/tests/kotlin/test_kotlin_big_numbers.rs

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
            val a = java.math.BigInteger("12345678901234567890")
            val b = java.math.BigInteger("987654321")
            __p((a.add(b).toString()).toString())
            __p((a.subtract(b).toString()).toString())
            __p((a.multiply(java.math.BigInteger("2")).toString()).toString())
        
__check("12345679888888888891\n12345677913580246769\n24691357802469135680")
}
