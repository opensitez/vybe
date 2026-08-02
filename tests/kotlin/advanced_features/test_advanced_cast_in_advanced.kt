// vybe-test: kotlin/advanced_features/test_advanced_cast_in_advanced
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val value: Any = 2
__check((value is Int).toString(), "true")
__check(((value as Int) + 3).toString(), "5") }
