// vybe-test: kotlin/default_arguments/test_default_arguments_method_defaults
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

class Counter {
            fun inc(value: Int = 1): Int = value
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Counter()
            __check((c.inc()).toString(), "1")
            __check((c.inc(3)).toString(), "3")
        }
