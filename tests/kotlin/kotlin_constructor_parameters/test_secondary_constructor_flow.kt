// vybe-test: kotlin/kotlin_constructor_parameters/test_secondary_constructor_flow
// origin: languages/kotlin/tests/kotlin/test_kotlin_constructor_parameters.rs

class Counter {
            val value: Int

            constructor(base: Int) {
                value = base
            }

            constructor() : this(0)

            fun isZero(): Boolean {
                return value == 0
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Counter().isZero()).toString(), "true")
            __check((Counter(4).isZero()).toString(), "false")
        }
