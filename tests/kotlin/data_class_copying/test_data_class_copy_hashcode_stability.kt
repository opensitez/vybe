// vybe-test: kotlin/data_class_copying/test_data_class_copy_hashcode_stability
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Entry(val key: String, val score: Int)
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
            val a = Entry("x", 1)
            val b = a.copy()
            __p((a.hashCode() == b.hashCode()).toString())
            __p((a == b).toString())
        
__check("true\ntrue")
}
