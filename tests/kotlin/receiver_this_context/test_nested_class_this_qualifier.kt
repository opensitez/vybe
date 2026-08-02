// vybe-test: kotlin/receiver_this_context/test_nested_class_this_qualifier
// origin: languages/kotlin/tests/kotlin/test_receiver_this_context.rs

class A {
            val name = "A"
            inner class B {
                fun call(): String = this@A.name
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((A().B().call()).toString(), "A")
        }
