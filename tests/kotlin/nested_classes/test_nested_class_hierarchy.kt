// vybe-test: kotlin/nested_classes/test_nested_class_hierarchy
// origin: languages/kotlin/tests/kotlin/test_nested_classes.rs

class Graph {
            class Node(val label: String)
            class Edge(val from: String, val to: String)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val node = Graph.Node("a")
            val edge = Graph.Edge("a", "b")
            __check((node.label).toString(), "a")
            __check((edge.from + edge.to).toString(), "ab")
        }
