// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_constructor_can_throw_and_fail
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Guard {
            val value: Int

            constructor(value: Int) {
                if (value < 0) {
                    throw Exception("invalid")
                }
                this.value = value
            }
        }

        fun main() {
            try {
                println(Guard(-2).value)
            } catch (e: Exception) {
                println("caught")
            }
        }

