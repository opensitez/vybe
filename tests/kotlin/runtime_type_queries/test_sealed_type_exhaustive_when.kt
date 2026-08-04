// vybe-test: kotlin/runtime_type_queries/test_sealed_type_exhaustive_when
// origin: languages/kotlin/tests/kotlin/test_runtime_type_queries.rs

sealed interface Node
        data class Leaf(val value: Int) : Node
        data class Branch(val left: Node, val right: Node) : Node

        fun classify(node: Node): String = when (node) {
            is Leaf -> "leaf"
            is Branch -> "branch"
        }

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
            val a = Leaf(1)
            val b = Branch(Leaf(2), Leaf(3))
            __p((classify(a)).toString())
            __p((classify(b)).toString())
        
__check("leaf\nbranch")
}
