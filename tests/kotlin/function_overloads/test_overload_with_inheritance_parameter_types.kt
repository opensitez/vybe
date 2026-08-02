// vybe-test: kotlin/function_overloads/test_overload_with_inheritance_parameter_types
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

open class Node
        class Child : Node()
        class Other : Node()
        fun visit(v: Node): String = "node"
        fun visit(v: Child): String = "child"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((visit(Child())).toString(), "child")
            __check((visit(Other())).toString(), "node")
        }
