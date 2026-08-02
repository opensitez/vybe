// vybe-test: kotlin/interfaces/test_interface_with_data_like_impl
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Shape { fun area(): Int }
class Square(val side: Int) : Shape { override fun area(): Int = side * side }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val s: Shape = Square(6)
__check((s.area()).toString(), "36") }
