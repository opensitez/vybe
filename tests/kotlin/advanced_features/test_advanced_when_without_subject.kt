// vybe-test: kotlin/advanced_features/test_advanced_when_without_subject
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val x = 4
val y = if (x < 0) 1 else 2
__check((y).toString(), "2") }
