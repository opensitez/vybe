// vybe-test: kotlin/object_declarations/test_object_access_in_companion_with_private_constructor
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

class Widget private constructor(val label: String) {
            companion object Holder {
                val shared = Maker
            }
        }

        object Maker : (String) -> Widget {
            override fun invoke(label: String): Widget = Widget(label)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Widget.Holder.shared("x").label).toString(), "x")
        }
