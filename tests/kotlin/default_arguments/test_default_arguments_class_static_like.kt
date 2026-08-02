// vybe-test: kotlin/default_arguments/test_default_arguments_class_static_like
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

class Counter {
            companion object {
                fun make(base: Int = 9): Int = base
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Counter.make()).toString(), "9")
            __check((Counter.make(2)).toString(), "2")
        }
