// vybe-test: kotlin/kotlin_initializer_blocks/test_init_order_simple
// origin: languages/kotlin/tests/kotlin/test_kotlin_initializer_blocks.rs

class Demo {
            init { __check(("a").toString(), "a") }
            init { __check(("b").toString(), "b") }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Demo()
        }
