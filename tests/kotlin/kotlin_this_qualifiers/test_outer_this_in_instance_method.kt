// vybe-test: kotlin/kotlin_this_qualifiers/test_outer_this_in_instance_method
// origin: languages/kotlin/tests/kotlin/test_kotlin_this_qualifiers.rs

class Outer {
            val name = "outer"

            inner class Inner {
                fun valueFromOuter(): String {
                    return this@Outer.name
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
            val out = Outer().Inner()
            __check((out.valueFromOuter()).toString(), "outer")
        }
