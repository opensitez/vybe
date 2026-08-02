// vybe-test: kotlin/constructor_chaining/test_constructor_private_constructor
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Secret private constructor(val v: Int) {
            companion object {
                fun create(v: Int) = Secret(v)
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Secret.create(9).v).toString(), "9")
        }
