// vybe-test: kotlin/classes/test_class_chain_methods
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Builder {
            var value: Int = 0
            fun set(value: Int): Builder {
                this.value = value
                return this
            }
            fun increment(step: Int): Builder {
                this.value += step
                return this
            }
            fun total(): Int {
                return this.value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Builder()
            val result = b.set(3).increment(4).total()
            __check((result).toString(), "7")
        }
