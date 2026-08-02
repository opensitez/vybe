// vybe-test: kotlin/kotlin_visibility_advanced/test_internal_visibility_within_module
// origin: languages/kotlin/tests/kotlin/test_kotlin_visibility_advanced.rs

internal class Box {
            fun payload() = 7
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box().payload()).toString(), "7")
        }
