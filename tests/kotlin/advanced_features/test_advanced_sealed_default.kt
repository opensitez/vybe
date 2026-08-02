// vybe-test: kotlin/advanced_features/test_advanced_sealed_default
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

sealed class Node { class A : Node()
class B : Node() }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val n: Node = Node.A()
if (n is Node.A) { __check(("a").toString(), "a") } }
