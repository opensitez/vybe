// vybe-test: kotlin/inner_classes/test_inner_class_returning_outer_reference
// origin: languages/kotlin/tests/kotlin/test_inner_classes.rs

class Logger {
            private var tag = "log"
            inner class Entry {
                fun marker(): Logger = this@Logger
            }

            fun tag(): String = tag
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val entry = Logger().Entry()
            __check((entry.marker().tag()).toString(), "log")
        }
