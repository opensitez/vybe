// vybe-test: kotlin/companion_objects/test_companion_object_factory_preserves_private_constructor_rules
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Token private constructor(val label: String) {
            companion object {
                fun create(prefix: String, suffix: Int): Token {
                    return Token(prefix + ":" + suffix.toString())
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Token.create("x", 9).label).toString(), "x:9")
        }
