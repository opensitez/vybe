// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_chain
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Rectangle {
            val width: Int
            val height: Int

            constructor(side: Int) : this(side, side) {
                __check(("square").toString(), "square")
            }

            constructor(width: Int, height: Int) {
                this.width = width
                this.height = height
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val square = Rectangle(3)
            __check((square.width).toString(), "3")
            __check((square.height).toString(), "3")
        }
