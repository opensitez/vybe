// vybe-test: kotlin/kotlin_visibility_advanced/test_local_visibility_on_constructor_property
// origin: languages/kotlin/tests/kotlin/test_kotlin_visibility_advanced.rs

class Holder(private val tag: String) {
            fun render() = tag
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Holder("x").render()).toString(), "x")
        }
