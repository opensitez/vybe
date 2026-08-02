// vybe-test: kotlin/classes/test_class_with_init_and_field
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Meter {
            val limit: Int
            init {
                __check(("init").toString(), "init")
            }
            constructor(value: Int) {
                this.limit = value
            }
            fun scale(): Int {
                return limit * 2
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = Meter(7)
            __check((m.scale()).toString(), "14")
        }
