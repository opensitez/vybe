// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_private_validation
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class PositiveCounter {
            val value: Int

            private constructor(value: Int) {
                this.value = value
            }

            constructor(raw: Int, valid: Boolean) : this(if (valid && raw > 0) raw else 0) {
                println("built")
            }
        }

        fun main() {
            println(PositiveCounter(-3, true).value)
            println(PositiveCounter(5, false).value)
            println(PositiveCounter(5, true).value)
        }

