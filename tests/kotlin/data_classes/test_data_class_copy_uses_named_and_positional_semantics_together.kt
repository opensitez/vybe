// vybe-test: kotlin/data_classes/test_data_class_copy_uses_named_and_positional_semantics_together
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Range(val start: Int, val end: Int)

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
            val base = Range(1, 10)
            val shifted = Range(0, base.end).copy(start = base.start + 1)
            __p((base.toString()).toString())
            __p((shifted.toString()).toString())
        
__check("Range(start=1, end=10)\nRange(start=2, end=10)")
}
