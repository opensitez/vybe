// vybe-test: kotlin/bitwise_operations/test_set_clear_and_toggle_idempotent
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
            val base = 0b1001
            val set2 = base or (1 shl 1)
            val clear2 = set2 and (1 shl 1).inv()
            val toggle = base xor (1 shl 2)
            val toggledBack = toggle xor (1 shl 2)
            __p((set2).toString())
            __p((clear2).toString())
            __p((toggle).toString())
            __p((toggledBack).toString())
        
__check("11\n9\n13\n9")
}
