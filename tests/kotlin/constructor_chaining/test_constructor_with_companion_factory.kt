// vybe-test: kotlin/constructor_chaining/test_constructor_with_companion_factory
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class PairNum private constructor(val a: Int, val b: Int) {
            companion object {
                fun of(a: Int) = PairNum(a, a + 1)
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = PairNum.of(2)
            __check((p.a + p.b).toString(), "5")
        }
