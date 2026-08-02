// vybe-test: kotlin/varargs/test_vararg_in_class_method
// origin: languages/kotlin/tests/kotlin/test_varargs.rs

class Collector {
            fun collect(vararg items: String): String = items.joinToString(",")
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Collector()
            __check((c.collect("x", "y")).toString(), "x,y")
        }
