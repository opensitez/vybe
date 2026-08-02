// vybe-test: kotlin/secondary_constructors/test_secondary_with_mutable_property
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class State {
            var total: Int

            constructor() {
                this.total = 0
            }

            constructor(init: Int) : this() {
                this.total = init
            }

            constructor(init: Int, add: Int) : this(init) {
                this.total = this.total + add
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = State(4, 5)
            __check((s.total).toString(), "9")
        }
