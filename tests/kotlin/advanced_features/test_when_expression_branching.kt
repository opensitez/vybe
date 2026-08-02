// vybe-test: kotlin/advanced_features/test_when_expression_branching
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

fun evaluate(x: Int) {
            when (x) {
                1 -> __check(("one").toString(), "one")
                2 -> __check(("two").toString(), "two")
                else -> __check(("other").toString(), "other")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            evaluate(1)
            evaluate(2)
            evaluate(99)
        }
