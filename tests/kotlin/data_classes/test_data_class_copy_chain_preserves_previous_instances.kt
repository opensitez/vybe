// vybe-test: kotlin/data_classes/test_data_class_copy_chain_preserves_previous_instances
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Node(val id: Int, val label: String)

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
            val a = Node(1, "a")
            val b = a.copy(label = "b")
            val c = b.copy(id = 3)
            __p((a.label).toString())
            __p((b.id).toString())
            __p((c.label).toString())
            __p((a == b).toString())
            __p((b == c).toString())
        
__check("a\n1\nb\nfalse\nfalse")
}
