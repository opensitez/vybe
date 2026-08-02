// vybe-test: kotlin/classes/test_class_inner_class_captures_outer_reference
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Outer(val prefix: String) {
            val marker = "!"

            inner class Inner(val value: String) {
                fun describe(): String {
                    return prefix + marker + value
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
            val outer = Outer("x")
            val inner = outer.Inner("y")
            __check((inner.describe()).toString(), "x!y")
        }
