// vybe-test: kotlin/variance/test_variance_nested_generics_contravariant_nested
// origin: languages/kotlin/tests/kotlin/test_variance.rs

interface Sink<in T> { fun consume(v: T) }
        open class Node
        class Leaf : Node()
        class NodeSink : Sink<Node> { override fun consume(v: Node) { __check((v::class.simpleName).toString(), "Leaf") } }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sink: Sink<Leaf> = NodeSink()
            sink.consume(Leaf())
        }
