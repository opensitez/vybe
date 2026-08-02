// vybe-test: kotlin/smart_casts/test_type_test_on_inherited_class_chain
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

open class Node
        open class Container : Node()
        class Boxed : Container()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Node = Boxed()
            __check((value is Node).toString(), "true")
            __check((value is Container).toString(), "true")
            __check((value is Boxed).toString(), "true")
        }
