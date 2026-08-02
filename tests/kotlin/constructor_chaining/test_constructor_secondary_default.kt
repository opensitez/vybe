// vybe-test: kotlin/constructor_chaining/test_constructor_secondary_default
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Box {
            val v: Int
            constructor(v: Int) { this.v = v }
            constructor() : this(3)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box().v).toString(), "3")
            __check((Box(5).v).toString(), "5")
        }
