// vybe-test: kotlin/secondary_constructors/test_secondary_with_init_printing
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Track {
            val value: Int
            init {
                println("init")
            }

            constructor() {
                this.value = 0
            }

            constructor(v: Int) : this() {
                this.value = v
            }
        }

        fun main() {
            Track()
            Track(3)
        }

