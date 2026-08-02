// vybe-test: kotlin/default_arguments/test_default_arguments_over_defaulted_class_methods
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

class Acc {
            fun add(a: Int, b: Int = 1): Int = a + b
            fun nested(label: String = "L"): String = label
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Acc()
            __check((a.add(3)).toString(), "4")
            __check((a.add(3, 4)).toString(), "7")
            __check((a.nested()).toString(), "L")
        }
