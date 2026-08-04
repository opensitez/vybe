// vybe-test: kotlin/equality_hashcode/test_list_indexof_uses_structural_equality
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Entry(val value: Int)

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
            val list = listOf(Entry(1), Entry(2))
            __p((list.indexOf(Entry(1))).toString())
            __p((list.indexOf(Entry(3))).toString())
        
__check("0\n-1")
}
