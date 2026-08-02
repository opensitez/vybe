// vybe-test: kotlin/imports/test_imports_with_no_usage_still_parses
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.math.max
        import kotlin.math.min
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((max(9, 3)).toString(), "9")
            __check((min(9, 3)).toString(), "3")
        }
