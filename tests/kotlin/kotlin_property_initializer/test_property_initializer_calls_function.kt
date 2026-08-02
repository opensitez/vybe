// vybe-test: kotlin/kotlin_property_initializer/test_property_initializer_calls_function
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_initializer.rs

var counter = 0

        fun next(): Int {
            counter = counter + 1
            return counter
        }

        class Node {
            val a: Int = next()
            val b: Int = next()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val n = Node()
            __check((n.a).toString(), "1")
            __check((n.b).toString(), "2")
            __check((counter).toString(), "2")
        }
