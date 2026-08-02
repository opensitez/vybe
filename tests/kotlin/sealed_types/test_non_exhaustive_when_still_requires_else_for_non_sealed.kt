// vybe-test: kotlin/sealed_types/test_non_exhaustive_when_still_requires_else_for_non_sealed
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Alpha {
            class A : Alpha()
        }

        open class Beta

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val alpha: Alpha = Alpha.A()
            __check((when (alpha) {
                is Alpha.A -> 1
            }).toString(), "1")
            val beta = Beta()
            __check((when (beta is Beta) {
                true -> 2
                false -> 3
            }).toString(), "2")
        }
