// vybe-test: kotlin/import_aliases/test_import_alias_local_override_of_import
// origin: languages/kotlin/tests/kotlin/test_import_aliases.rs

import kotlin.math.sqrt as squareRoot

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun squareRoot(x: Int): Int = x * x
            val f: (Double) -> Double = kotlin.math::sqrt
            __check((squareRoot(3)).toString(), "9")
            __check((f(4.0).toInt()).toString(), "2")
        }
