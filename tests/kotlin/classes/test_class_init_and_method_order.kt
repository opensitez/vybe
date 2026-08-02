// vybe-test: kotlin/classes/test_class_init_and_method_order
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Logger { init { __check(("start").toString(), "start") }
fun value(): Int = 5 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Logger().value()).toString(), "5") }
