// vybe-test: kotlin/generics/test_generic_bound_charsequence_length
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T : CharSequence> totalLength(left: T, right: T): Int {
            return left.length + right.length
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((totalLength("ab", "xyz")).toString(), "5")
            __check((totalLength(StringBuilder("k"), StringBuilder("on"))).toString(), "3")
        }
