// vybe-test: kotlin/local_functions/test_local_function_inside_class_method
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

class Engine {
            fun execute(base: Int): Int {
                fun bump(v: Int): Int = v + 1
                return bump(base)
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Engine().execute(7)).toString(), "8")
        }
