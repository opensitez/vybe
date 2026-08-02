// vybe-test: kotlin/advanced_features/test_advanced_extension_in_advanced
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

class Holder(val value: Int)
fun Holder.double() = value * 2
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Holder(4).double()).toString(), "8") }
