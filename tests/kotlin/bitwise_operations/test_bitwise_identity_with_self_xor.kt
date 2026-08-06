// vybe-test: kotlin/bitwise_operations/test_bitwise_identity_with_self_xor
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
            val values = listOf(0, 1, 2, 3, 255)
            val unchanged = values.map { it xor it }
            val back = values.map { (it xor 0) xor it }
            __p((unchanged.joinToString(",")).toString())
            __p((back.joinToString(",")).toString())
        
// Real Kotlin agrees: `(it xor 0) xor it` is `it xor it` = 0 for every
// value — a round-trip back to the input needs a nonzero key on both
// sides, which this body never uses.
__check("0,0,0,0,0\n0,0,0,0,0")
}
