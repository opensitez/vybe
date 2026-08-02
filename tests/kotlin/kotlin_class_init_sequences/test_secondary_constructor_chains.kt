// vybe-test: kotlin/kotlin_class_init_sequences/test_secondary_constructor_chains
// origin: languages/kotlin/tests/kotlin/test_kotlin_class_init_sequences.rs

class Node {
            val x: Int
            constructor(a: Int) {
                x = a
            }
            constructor(a: Int, b: Int) : this(a + b)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n = Node(2, 3)
            __check((n.x).toString(), "5")
        }
