// vybe-test: kotlin/constructor_chaining/test_constructor_chained_via_this
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class ThisChain {
            val v: Int
            constructor() : this(8)
            constructor(v: Int) { this.v = v }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((ThisChain().v).toString(), "8")
        }
