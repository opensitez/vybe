// vybe-test: kotlin/sealed_types/test_nested_sealed_subclasses_keep_disjoint_branches
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Shape {
            sealed class Circle : Shape() {
                class Small : Circle()
                class Large : Circle()
            }

            class Square : Shape()
        }

        fun area(shape: Shape): String {
            return when (shape) {
                is Shape.Circle.Small -> "small"
                is Shape.Circle.Large -> "large"
                is Shape.Square -> "square"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((area(Shape.Circle.Small())).toString(), "small")
            __check((area(Shape.Circle.Large())).toString(), "large")
            __check((area(Shape.Square())).toString(), "square")
        }
