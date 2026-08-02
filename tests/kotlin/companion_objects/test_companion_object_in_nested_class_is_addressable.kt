// vybe-test: kotlin/companion_objects/test_companion_object_in_nested_class_is_addressable
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Holder {
            class Nested {
                companion object {
                    fun label(value: Int): String = "id:" + value
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
            __check((Holder.Nested.label(7)).toString(), "id:7")
        }
