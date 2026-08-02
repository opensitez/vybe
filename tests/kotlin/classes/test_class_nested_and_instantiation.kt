// vybe-test: kotlin/classes/test_class_nested_and_instantiation
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Outer {
            class Inner {
                fun ping(): String {
                    return "pong"
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
            val i = Outer.Inner()
            __check((i.ping()).toString(), "pong")
        }
