// vybe-test: kotlin/companion_objects/test_companion_with_extension_style_call_site
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Labeler {
            companion object {
                fun from(prefix: String, value: Int): String = prefix + value.toString()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Labeler.from("v", 4)).toString(), "v4")
        }
