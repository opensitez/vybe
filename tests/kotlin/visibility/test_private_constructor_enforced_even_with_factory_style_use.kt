// vybe-test: kotlin/visibility/test_private_constructor_enforced_even_with_factory_style_use
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Config private constructor(val value: Int) {
            companion object {
                fun create(value: Int): Config = Config(value)
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Config.create(3).value).toString(), "3")
        }
