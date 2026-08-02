// vybe-test: kotlin/advanced_features/test_advanced_generic_like
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

open class Box(val item: Int) { fun value(): Int = item }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val b = Box(7)
__check((b.value()).toString(), "7") }
