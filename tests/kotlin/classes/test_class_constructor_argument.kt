// vybe-test: kotlin/classes/test_class_constructor_argument
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Config(val timeout: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Config(30)
            __check((c.timeout).toString(), "30")
        }
