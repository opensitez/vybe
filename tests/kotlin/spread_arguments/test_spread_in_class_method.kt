// vybe-test: kotlin/spread_arguments/test_spread_in_class_method
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

class Acc {
            fun join(vararg values: Int): Int = values.size
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Acc()
            __check((a.join(1,2)).toString(), "2")
        }
