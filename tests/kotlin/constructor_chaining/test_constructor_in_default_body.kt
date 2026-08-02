// vybe-test: kotlin/constructor_chaining/test_constructor_in_default_body
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class DefaultBody {
            val value: Int
            constructor(v: Int = 1) { value = v }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((DefaultBody().value).toString(), "1")
        }
