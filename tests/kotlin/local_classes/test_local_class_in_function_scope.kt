// vybe-test: kotlin/local_classes/test_local_class_in_function_scope
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun wrap(v: Int): Int {
            class Local {
                fun value() = v + 1
            }
            return Local().value()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((wrap(3)).toString(), "4")
        }
