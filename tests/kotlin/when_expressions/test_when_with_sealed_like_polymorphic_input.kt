// vybe-test: kotlin/when_expressions/test_when_with_sealed_like_polymorphic_input
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

sealed class Node {
            class A(val value: Int) : Node()
            class B(val value: String) : Node()
            class C : Node()
        }

        fun render(node: Node): String {
            return when (node) {
                is Node.A -> "A:" + node.value
                is Node.B -> "B:" + node.value
                is Node.C -> "C"
            }
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
            __p((render(Node.A(9))).toString())
            __p((render(Node.B("x"))).toString())
            __p((render(Node.C())).toString())
        
__check("A:9\nB:x\nC")
}
