// vybe-test: kotlin/classes/test_nested_class_with_reference
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Outer { class Inner { fun read(): Int = 3 } }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val i = Outer.Inner()
__check((i.read()).toString(), "3") }
