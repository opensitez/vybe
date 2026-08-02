// vybe-test: kotlin/kotlin_package_aliases/test_import_alias_for_function
// origin: languages/kotlin/tests/kotlin/test_kotlin_package_aliases.rs

import kotlin.math.abs as kotlinAbs

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((kotlinAbs(-7)).toString(), "7")
        }
