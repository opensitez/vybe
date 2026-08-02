// vybe-test: kotlin/imports/test_import_star_import_used
// origin: languages/kotlin/tests/kotlin/test_imports.rs

import kotlin.math.*
        import kotlin.math.PI
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((PI.toInt()).toString(), "3")
            __check((round(3.4)).toString(), "3")
        }
