// vybe-test: kotlin/kotlin_package_aliases/test_import_alias_keeps_call_site_clear
// origin: languages/kotlin/tests/kotlin/test_kotlin_package_aliases.rs

import kotlin.math.max as m

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((m(10, 2)).toString(), "10")
        }
