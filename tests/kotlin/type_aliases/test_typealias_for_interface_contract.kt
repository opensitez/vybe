// vybe-test: kotlin/type_aliases/test_typealias_for_interface_contract
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

interface Handler {
            fun run(value: Int): Int
        }

        typealias IncrHandler = Handler

        object Adder : IncrHandler {
            override fun run(value: Int): Int = value + 1
        }

        fun apply(handler: IncrHandler, value: Int): Int = handler.run(value)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((apply(Adder, 8)).toString(), "9")
        }
