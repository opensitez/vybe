// vybe-test: kotlin/generics/test_generic_typealias_and_alias_target
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pair = Holder("left", "right")
            __check((pair.parts()).toString(), "left:right")
        }
