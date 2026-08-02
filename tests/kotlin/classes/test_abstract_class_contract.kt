// vybe-test: kotlin/classes/test_abstract_class_contract
// origin: languages/kotlin/tests/kotlin/test_classes.rs

abstract class Node { abstract fun id(): Int }
class Leaf : Node() { override fun id(): Int = 9 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val n: Node = Leaf()
__check((n.id()).toString(), "9") }
