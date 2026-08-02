// vybe-test: kotlin/kotlin_property_initializer/test_initializer_with_expression
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_initializer.rs

class Meter {
            val label = "x" + 1.toString() + "y"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Meter().label).toString(), "x1y")
        }
