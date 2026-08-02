// vybe-test: kotlin/classes/test_class_with_abstract_implementation
// origin: languages/kotlin/tests/kotlin/test_classes.rs

abstract class Worker {
            abstract fun work(): Int
        }

        class Coder : Worker() {
            override fun work(): Int {
                return 9
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val w: Worker = Coder()
            __check((w.work()).toString(), "9")
        }
