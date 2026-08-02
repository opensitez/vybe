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

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((render(Node.A(9))).toString(), "A:9")
            __check((render(Node.B("x"))).toString(), "B:x")
            __check((render(Node.C())).toString(), "C")
        }
