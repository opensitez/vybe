// vybe-test: kotlin/inner_classes/test_inner_class_from_method
// origin: languages/kotlin/tests/kotlin/test_inner_classes.rs

class Builder {
            private val base = 2
            inner class Worker(val factor: Int) {
                fun total(): Int = base * factor
            }

            fun make(): Int = Worker(4).total()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Builder().make()).toString(), "8")
        }
