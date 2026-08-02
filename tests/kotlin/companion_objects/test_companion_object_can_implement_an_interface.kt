// vybe-test: kotlin/companion_objects/test_companion_object_can_implement_an_interface
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

interface Stamp {
            fun stamp(value: String): String
        }

        class Tagger {
            companion object : Stamp {
                override fun stamp(value: String): String = "tagged-" + value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Tagger.stamp("a")).toString(), "tagged-a")
            __check((Tagger.stamp("b")).toString(), "tagged-b")
        }
