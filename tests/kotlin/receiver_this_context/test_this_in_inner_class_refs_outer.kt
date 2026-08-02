// vybe-test: kotlin/receiver_this_context/test_this_in_inner_class_refs_outer
// origin: languages/kotlin/tests/kotlin/test_receiver_this_context.rs

class Outer(val outerLabel: String) {
            inner class Inner {
                fun full(): String = this@Outer.outerLabel + ":inner"
            }

            fun probe(): String = Inner().full()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Outer("root").probe()).toString(), "root:inner")
        }
