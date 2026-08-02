// vybe-test: kotlin/constructor_chaining/test_constructor_secondary_chain_to_primary
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Chain {
            val v: Int
            constructor() { this(7) }
            constructor(v: Int) { this.v = v }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Chain().v).toString(), "7")
        }
