// vybe-test: kotlin/constructor_chaining/test_constructor_two_secondary_calls
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Layer {
            val a: Int
            val b: Int
            constructor(a: Int) { this.a = a
this.b = a }
            constructor(a: Int, b: Int) : this(a + b)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val l = Layer(2, 3)
            __check((l.a).toString(), "5")
            __check((l.b).toString(), "5")
        }
