// vybe-test: kotlin/visibility/test_factory_function_preserves_private_constructor_access_control
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Box private constructor(val value: String) {
            companion object {
                fun from(value: String): Box = Box(value)
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box.from("x").value).toString(), "x")
        }
