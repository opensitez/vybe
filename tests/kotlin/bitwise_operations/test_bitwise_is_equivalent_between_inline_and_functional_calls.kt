// vybe-test: kotlin/bitwise_operations/test_bitwise_is_equivalent_between_inline_and_functional_calls
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
            val base = 0b11011001
            val andResult = base and 0x0F
            val andAlt = kotlin.math.floor(base.toDouble()).toInt() and 0x0F
            __p((andResult).toString())
            __p((andAlt).toString())
            val invAnd = base and (1 shl 4).inv()
            __p((invAnd).toString())
        
__check("25\n25\n201")
}
