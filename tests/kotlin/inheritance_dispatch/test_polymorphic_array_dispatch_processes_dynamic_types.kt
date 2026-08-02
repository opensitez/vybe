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

        fun main() {
            val nodes: Array<Node> = arrayOf(Node(), Leaf(), Branch())
            println(summarize(nodes))
        }

