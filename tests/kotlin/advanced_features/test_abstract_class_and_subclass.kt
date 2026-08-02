// vybe-test: kotlin/advanced_features/test_abstract_class_and_subclass
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

abstract class Shape {
            abstract fun area(): Int
            fun describe() {
                __check(("Shape area is " + area()).toString(), "Shape area is 25")
            }
        }

        class Square(val side: Int) : Shape() {
            override fun area(): Int = side * side
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = Square(5)
            s.describe()
        }
