// vybe-test: kotlin/kotlin_nested_scope_functions/test_nested_with_and_apply_combo
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_scope_functions.rs

class Box {
            var value = 1
            fun bump() { value += 1 }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box().apply {
                bump()
                value += 2
            }.run {
                "${'$'}value"
            }
            __check((b).toString(), "4")
        }
