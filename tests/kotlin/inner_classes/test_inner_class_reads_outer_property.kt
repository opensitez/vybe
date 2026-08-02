// vybe-test: kotlin/inner_classes/test_inner_class_reads_outer_property
// origin: languages/kotlin/tests/kotlin/test_inner_classes.rs

class Outer(val base: Int) {
            inner class Inner(val delta: Int)
            fun make(): Int = Inner(3).delta + base
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Outer(10).make()).toString(), "13")
        }
