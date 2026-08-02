// vybe-test: kotlin/kotlin_package_aliases/test_package_import_star_behavior
// origin: languages/kotlin/tests/kotlin/test_kotlin_package_aliases.rs

import kotlin.math.*

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sin(0.0) == 0.0).toString(), "true")
            __check((cos(0.0) == 1.0).toString(), "true")
        }
