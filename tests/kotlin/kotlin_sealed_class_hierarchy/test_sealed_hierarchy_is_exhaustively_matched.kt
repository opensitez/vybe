// vybe-test: kotlin/kotlin_sealed_class_hierarchy/test_sealed_hierarchy_is_exhaustively_matched
// origin: languages/kotlin/tests/kotlin/test_kotlin_sealed_class_hierarchy.rs

sealed class Node {
            data class Value(val n: Int) : Node()
            data class Negate(val child: Node) : Node()
            data class Sum(val left: Node, val right: Node) : Node()
        }

        fun eval(node: Node): Int = when (node) {
            is Node.Value -> node.n
            is Node.Negate -> -eval(node.child)
            is Node.Sum -> eval(node.left) + eval(node.right)
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
            val expr = Node.Sum(Node.Value(3), Node.Negate(Node.Value(2)))
            __p((eval(expr)).toString())
        
__check("1")
}
