// vybe-test: kotlin/classes/test_class_with_companion_counter
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Factory { companion object { var index = 0
fun next(): Int { index += 1
return index } } }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Factory.next()).toString(), "1")
__check((Factory.next()).toString(), "2") }
