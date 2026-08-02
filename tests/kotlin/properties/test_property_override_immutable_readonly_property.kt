// vybe-test: kotlin/properties/test_property_override_immutable_readonly_property
// origin: languages/kotlin/tests/kotlin/test_properties.rs

open class Node {
            open val label: String = "base"
        }

        class Leaf : Node() {
            override val label: String = "leaf"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val node: Node = Leaf()
            __check((node.label).toString(), "leaf")
        }
