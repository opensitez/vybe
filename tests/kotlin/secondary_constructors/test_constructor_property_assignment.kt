// vybe-test: kotlin/secondary_constructors/test_constructor_property_assignment
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Rectangle {
            val width: Int
            val height: Int

            constructor(width: Int, height: Int) {
                this.width = width
                this.height = height
            }

            constructor(size: Int) : this(size, size) {
                __check(("square").toString(), "square")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r1 = Rectangle(2, 3)
            val r2 = Rectangle(4)
            __check((r1.width).toString(), "2")
            __check((r1.height).toString(), "3")
            __check((r2.width).toString(), "4")
            __check((r2.height).toString(), "4")
        }
