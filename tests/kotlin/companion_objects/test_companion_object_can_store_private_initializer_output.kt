// vybe-test: kotlin/companion_objects/test_companion_object_can_store_private_initializer_output
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Builder {
            companion object {
                private const val prefix = "id:"
                val marker: String

                init {
                    marker = prefix + "1"
                }

                fun label(value: Int): String {
                    return marker + value.toString()
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
            __check((Builder.label(4)).toString(), "id:14")
        }
