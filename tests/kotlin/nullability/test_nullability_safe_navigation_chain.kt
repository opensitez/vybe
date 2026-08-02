// vybe-test: kotlin/nullability/test_nullability_safe_navigation_chain
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

class Box(val value: Int)
class Wrapper(val box: Box?)
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val wrapped: Wrapper? = Wrapper(Box(4))
val fallback = wrapped?.box?.value ?: -1
__check((fallback).toString(), "4") }
