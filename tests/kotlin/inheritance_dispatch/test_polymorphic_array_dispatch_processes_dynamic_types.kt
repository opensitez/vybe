// vybe-test: kotlin/inheritance_dispatch/test_polymorphic_array_dispatch_processes_dynamic_types
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Node {
            open fun kind(): String = "node"
        }

        class Leaf : Node() {
            override fun kind(): String = "leaf"
        }

        class Branch : Node() {
            override fun kind(): String = "branch"
        }

        fun summarize(nodes: Array<Node>): String {
            var value = ""
            for (node in nodes) {
                value += node.kind()
                value += ";"
            }
            return value
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
            val nodes: Array<Node> = arrayOf(Node(), Leaf(), Branch())
            __p((summarize(nodes)).toString())
        
__check("node;leaf;branch;")
}
