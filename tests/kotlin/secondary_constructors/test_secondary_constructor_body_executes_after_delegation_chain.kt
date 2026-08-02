// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_body_executes_after_delegation_chain
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

var trace = ""

        class Chain {
            var value: Int

            constructor() {
                trace += "root;"
                value = 1
            }

            constructor(value: Int) : this() {
                trace += "inner;"
                this.value = value
            }

            constructor(value: Int, extra: Int) : this(value) {
                trace += "leaf;"
                this.value = value + extra
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Chain(2, 3)
            __check((trace).toString(), "root;inner;leaf;")
            __check((c.value).toString(), "5")
        }
