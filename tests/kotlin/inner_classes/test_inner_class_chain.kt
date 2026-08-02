// vybe-test: kotlin/inner_classes/test_inner_class_chain
// origin: languages/kotlin/tests/kotlin/test_inner_classes.rs

class Network {
            val root = "R"
            inner class Segment {
                inner class Node(val id: Int) {
                    fun label(): String = root + id
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val node = Network().Segment().Node(7)
            __check((node.label()).toString(), "R7")
        }
