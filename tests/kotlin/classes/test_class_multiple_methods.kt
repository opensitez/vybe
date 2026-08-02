// vybe-test: kotlin/classes/test_class_multiple_methods
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Calc(val base: Int) {
            fun add(x: Int): Int = base + x
            fun mul(x: Int): Int = base * x
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Calc(5)
            __check((c.add(3)).toString(), "8")
            __check((c.mul(4)).toString(), "20")
        }
