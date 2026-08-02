// vybe-test: kotlin/infix/test_infix_chainable_custom
// origin: languages/kotlin/tests/kotlin/test_infix.rs

class Box(val value: Int) { infix fun plus(other: Box): Int = value + other.value
infix fun minus(other: Box): Int = value - other.value }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = Box(9)
val b = Box(3)
__check((a plus b).toString(), "12")
__check((a minus b).toString(), "6") }
