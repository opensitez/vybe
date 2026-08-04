// vybe-test: kotlin/collections_set/test_mutable_set_plus_assign_with_elements
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

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
            val base = mutableSetOf(1, 2)
            val snapshot = base.toSet()
            base += setOf(2, 3, 4)
            __p((base.size).toString())
            __p((snapshot.size).toString())
            __p((base.contains(4)).toString())
        
__check("4\n2\ntrue")
}
