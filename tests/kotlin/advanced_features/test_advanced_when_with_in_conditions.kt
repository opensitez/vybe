// vybe-test: kotlin/advanced_features/test_advanced_when_with_in_conditions
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val score = 77
            val label = when (score) {
                in 90..100 -> "A"
                in 80..89 -> "B"
                in 70..79 -> "C"
                else -> "F"
            }
            __check((label).toString(), "C")
        }
