// vybe-test: kotlin/companion_objects/test_companion_object_implements_function_type
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Prefixer {
            companion object : (String) -> String {
                override fun invoke(value: String): String = ">> " + value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: (String) -> String = Prefixer.Companion
            __check((value("a")).toString(), ">> a")
            __check((Prefixer.Companion("b")).toString(), ">> b")
        }
