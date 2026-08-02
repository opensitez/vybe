// vybe-test: kotlin/member_references/test_reference_to_math_abs_function
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

import kotlin.math.abs
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = ::abs
            __check((f(-10)).toString(), "10")
        }
