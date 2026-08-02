// vybe-test: kotlin/secondary_constructors/test_constructor_three_levels
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Scale {
            val unit: Int

            constructor() {
                this.unit = 1
            }

            constructor(value: Int) : this() {
                println("scaled")
            }

            constructor(value: Int, factor: Int) : this(value) {
                println(value * factor)
            }
        }

        fun main() {
            Scale()
            Scale(4)
            Scale(5, 2)
        }

