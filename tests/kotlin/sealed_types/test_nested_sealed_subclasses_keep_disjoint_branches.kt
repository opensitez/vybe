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

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __p((area(Shape.Circle.Small())).toString())
            __p((area(Shape.Circle.Large())).toString())
            __p((area(Shape.Square())).toString())
        
__check("small\nlarge\nsquare")
}
