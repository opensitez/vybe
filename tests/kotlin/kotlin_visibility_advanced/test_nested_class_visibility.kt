// vybe-test: kotlin/kotlin_visibility_advanced/test_nested_class_visibility
// origin: languages/kotlin/tests/kotlin/test_kotlin_visibility_advanced.rs

class Outer {
            private val secret = "open"
            inner class Inner {
                fun reveal() = secret
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Outer().Inner().reveal()).toString(), "open")
        }
