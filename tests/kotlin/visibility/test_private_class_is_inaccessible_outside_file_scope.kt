// vybe-test: kotlin/visibility/test_private_class_is_inaccessible_outside_file_scope
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

private class Local {
            fun value(): String = "inner"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Local().value()).toString(), "inner")
        }
