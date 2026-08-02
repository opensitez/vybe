// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_delegation_argument_evaluation_order
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

var order = ""

        fun mark(value: String): Int {
            order += value
            return if (value == "left") 1 else 2
        }

        class Probe {
            val first: Int
            val second: Int

            constructor(first: Int, second: Int) {
                this.first = first
                this.second = second
            }

            constructor(value: Int) : this(mark("left"), mark("right") + value) {
                // body intentionally empty
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val probe = Probe(3)
            __check((order).toString(), "leftright")
            __check((probe.first).toString(), "1")
            __check((probe.second).toString(), "5")
        }
