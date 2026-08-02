// vybe-test: kotlin/advanced_features/test_advanced_nested_conditional
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

fun score(x: Int): String { return if (x > 10) "high" else if (x > 5) "mid" else "low" }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((score(11)).toString(), "high")
__check((score(3)).toString(), "low") }
