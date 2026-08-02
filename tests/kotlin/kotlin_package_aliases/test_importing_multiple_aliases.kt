// vybe-test: kotlin/kotlin_package_aliases/test_importing_multiple_aliases
// origin: languages/kotlin/tests/kotlin/test_kotlin_package_aliases.rs

import kotlin.math.max as takeMax
        import kotlin.math.min as takeMin

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((takeMax(3, 7)).toString(), "7")
            __check((takeMin(3, 7)).toString(), "3")
        }
