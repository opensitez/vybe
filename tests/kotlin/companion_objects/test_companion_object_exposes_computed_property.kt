// vybe-test: kotlin/companion_objects/test_companion_object_exposes_computed_property
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Converter {
            companion object {
                const val base = 100
                val scaled: Int
                    get() = base * 2
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Converter.base).toString(), "100")
            __check((Converter.scaled).toString(), "200")
        }
