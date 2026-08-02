// vybe-test: kotlin/kotlin_initializer_blocks/test_init_inheritance_order
// origin: languages/kotlin/tests/kotlin/test_kotlin_initializer_blocks.rs

open class Parent {
            init { __check(("p").toString(), "p") }
        }

        class Child : Parent() {
            init { __check(("c").toString(), "c") }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Child()
        }
