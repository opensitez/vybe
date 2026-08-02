// vybe-test: kotlin/reified_generics/test_reified_generic_identity
// origin: languages/kotlin/tests/kotlin/test_reified_generics.rs

inline fun <reified T> sameType(a: T, b: T): String = if (a::class == b::class) "same" else "diff"

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sameType(1, 2)).toString(), "same")
            __check((sameType("a", "b")).toString(), "same")
        }
