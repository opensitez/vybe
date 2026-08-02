// vybe-test: kotlin/classes/test_class_with_constructor_sharing
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class PairNode {
            val left: Int
            val right: Int

            constructor(left: Int, right: Int) {
                this.left = left
                this.right = right
            }

            constructor(value: Int) : this(value, value) {
                __check(("copy").toString(), "copy")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p1 = PairNode(4)
            val p2 = PairNode(1, 3)
            __check((p1.left).toString(), "4")
            __check((p1.right).toString(), "4")
            __check((p2.left).toString(), "1")
            __check((p2.right).toString(), "3")
        }
