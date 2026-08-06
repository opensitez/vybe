// vybe-test: kotlin/bitwise_operations/test_bitwise_precedence_with_arithmetic
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
            __p((1 or 2 + 4 and 8).toString())
            __p(((1 or 2) + (4 and 8)).toString())
            __p((2 shl 3 + 1).toString())
            __p((2 shl (3 + 1)).toString())
        
// Real Kotlin agrees: additive binds tighter than the named infix ops, and
// same-level infix chains are LEFT-associative — `1 or 2 + 4 and 8` is
// `(1 or 6) and 8` = 0, and `2 shl 3 + 1` is `2 shl 4` = 32.
__check("0\n3\n32\n32")
}
